using System.Text.RegularExpressions;
using OutlookMcp.Application.Configuration;
using OutlookMcp.Application.Services;
using OutlookMcp.Contracts;

namespace OutlookMcp.Application.WritingStyle;

public sealed partial class AuthoredTextExtractor
{
    private readonly EmailBodyCleaner _bodyCleaner;
    private readonly IReadOnlyList<string> _signatureMarkers;
    private readonly IReadOnlyList<string> _disclaimerMarkers;

    public AuthoredTextExtractor(EmailBodyCleaner bodyCleaner, OutlookMcpOptions? options = null)
    {
        _bodyCleaner = bodyCleaner;
        _signatureMarkers = options?.WritingStyle.AdditionalSignatureMarkers ?? [];
        _disclaimerMarkers = options?.WritingStyle.AdditionalDisclaimerMarkers ?? [];
    }

    [GeneratedRegex(@"(?im)^\s*(?:From|Saatja|Von|От):\s+.+$|^\s*-{2,}\s*(?:Original Message|Forwarded message|Edastatud kiri).*$|^\s*On .+ wrote:\s*$|^\s*.+ kirjutas:\s*$|^\s*>+")]
    private static partial Regex QuoteBoundary();

    [GeneratedRegex(@"(?im)^\s*(?:Lugupidamisega|Parimate soovidega|Tervitades|Regards|Best regards|Kind regards)[,!]?\s*$")]
    private static partial Regex SignatureBoundary();

    [GeneratedRegex(@"(?im)^.*(?:confidential|konfidentsiaal|käesolev e-kiri|this e-mail and any attachments|intended recipient).*$")]
    private static partial Regex DisclaimerBoundary();

    [GeneratedRegex(@"(?m)^\s*(?:[-*•]|\d+[.)])\s+")]
    private static partial Regex ListItem();

    [GeneratedRegex(@"(?im)^\s*(Tere(?:,?\s+[^!\n]+)?!?|Tervist!?|Hei!?|Hello!?|Dear\s+[^\n]+)[ \t]*$")]
    private static partial Regex GreetingLine();

    public AuthoredTextExtractionDto Extract(string? plainText, string? html)
    {
        var clean = _bodyCleaner.Clean(plainText, html).Complete;
        if (string.IsNullOrWhiteSpace(clean)) return Empty("body_unavailable", "No readable plain-text or HTML body was available.");

        var quoteMatch = QuoteBoundary().Match(clean);
        var authoredAndSignature = quoteMatch.Success ? clean[..quoteMatch.Index].Trim() : clean;
        var quoted = quoteMatch.Success ? clean[quoteMatch.Index..].Trim() : string.Empty;

        var disclaimerMatch = DisclaimerBoundary().Match(authoredAndSignature);
        var disclaimerIndex = FirstBoundary(disclaimerMatch.Success ? disclaimerMatch.Index : null, authoredAndSignature, _disclaimerMarkers);
        var disclaimer = disclaimerIndex is not null ? authoredAndSignature[disclaimerIndex.Value..].Trim() : string.Empty;
        if (disclaimerIndex is not null) authoredAndSignature = authoredAndSignature[..disclaimerIndex.Value].Trim();

        var signatureMatch = SignatureBoundary().Match(authoredAndSignature);
        var signatureIndex = FirstBoundary(signatureMatch.Success ? signatureMatch.Index : null, authoredAndSignature, _signatureMarkers);
        var signature = signatureIndex is not null ? authoredAndSignature[signatureIndex.Value..].Trim() : string.Empty;
        var authored = signatureIndex is not null ? authoredAndSignature[..signatureIndex.Value].Trim() : authoredAndSignature.Trim();
        var closing = signatureMatch.Success ? signatureMatch.Value.Trim() : signatureIndex is not null ? FirstLine(signature) : null;
        var greeting = GreetingLine().Match(authored).Success ? GreetingLine().Match(authored).Value.Trim() : null;

        var confidence = 0.95;
        if (!quoteMatch.Success && clean.Contains('@') && clean.Length > 2_000) confidence -= 0.12;
        if (signature.Length == 0 && Regex.IsMatch(authored, @"(?m)^\+?\d[\d\s-]{6,}$", RegexOptions.CultureInvariant)) confidence -= 0.08;
        if (authored.Length > 15_000) confidence -= 0.12;
        if (authored.Length < 2) confidence = 0.2;
        confidence = Math.Clamp(confidence, 0.05, 1.0);

        var paragraphs = authored.Split("\n\n", StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Length;
        var status = string.IsNullOrWhiteSpace(authored) ? "no_authored_text_detected" : "successfully_processed";
        return new AuthoredTextExtractionDto(clean, authored, quoted, signature, disclaimer, confidence,
            quoteMatch.Success ? "reply_boundary_and_signature_heuristics" : "signature_and_disclaimer_heuristics",
            status, status == "successfully_processed" ? null : "Only quoted, signature, or disclaimer content was detected.",
            greeting, closing, paragraphs, ListItem().Matches(authored).Count, authored.Count(character => character == '?'));
    }

    private static AuthoredTextExtractionDto Empty(string status, string reason) => new(
        string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, 0, "no_body", status, reason, null, null, 0, 0, 0);

    private static int? FirstBoundary(int? detected, string text, IReadOnlyList<string> markers)
    {
        var result = detected;
        foreach (var marker in markers)
        {
            var index = text.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (index >= 0 && (result is null || index < result)) result = index;
        }
        return result;
    }

    private static string FirstLine(string value)
    {
        var index = value.IndexOf('\n');
        return (index < 0 ? value : value[..index]).Trim();
    }
}

public sealed class CommunicationIntentClassifier
{
    private static readonly (string Intent, string[] Terms)[] Rules =
    [
        ("scheduling", ["kohtume", "kell", "sobiks", "meeting", "schedule", "aeg"]),
        ("requesting_missing_information", ["palun täpsusta", "palun saata", "puudu", "vajame", "please provide"]),
        ("requesting_approval", ["palun kinnita", "kooskõlast", "approval", "approve"]),
        ("answering_technical_question", ["vastuseks", "lahendus", "tehnilis", "according to", "answer"]),
        ("clarifying_design_scope", ["töömaht", "scope", "projekteerimise piir", "ei kuulu"]),
        ("explaining_responsibility", ["vastutus", "vastutab", "responsibility", "eule osa"]),
        ("giving_status_update", ["seis", "staatus", "valmis", "progress", "hetkel"]),
        ("confirming_agreement", ["kinnitan", "kokku lepitud", "sobib", "confirmed", "agreed"]),
        ("disagreeing_or_correcting", ["ei ole õige", "parandus", "täpsustan", "however", "incorrect"]),
        ("sending_deliverables", ["saadan", "manuses", "edastan", "attached", "deliverable"]),
        ("following_up", ["tuletan meelde", "kas olete jõudnud", "follow up", "meeldetuletus"]),
        ("acknowledging_receipt", ["kätte saadud", "aitäh, sain", "received", "tänan"]),
        ("explaining_delay", ["viibib", "hilineb", "delay", "kahjuks ei jõua"]),
        ("providing_recommendation", ["soovitame", "minu soovitus", "recommend"]),
        ("asking_for_decision", ["palun otsust", "kumb variant", "decision", "valige"])
    ];

    public string Classify(string? subject, string? text)
    {
        var value = (subject + " " + text).ToLowerInvariant();
        return Rules.Select((rule, index) => (rule.Intent, Score: rule.Terms.Count(term => value.Contains(term, StringComparison.Ordinal)), Index: index))
            .Where(value => value.Score > 0).OrderByDescending(value => value.Score).ThenBy(value => value.Index)
            .Select(value => value.Intent).FirstOrDefault() ?? "general_correspondence";
    }
}
