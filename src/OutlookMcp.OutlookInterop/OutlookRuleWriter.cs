using System.Collections;
using System.Reflection;
using Microsoft.Extensions.Logging;
using OutlookMcp.Contracts;

namespace OutlookMcp.OutlookInterop;

internal sealed class OutlookRuleWriter(ILogger logger)
{
    private const int ReceiveRule = 0;

    public object CreateAndSave(object rulesObject, object destinationObject, CreateFolderRuleRequest request)
    {
        dynamic rules = rulesObject;
        dynamic? createdRule = null;
        var originalCount = (int)rules.Count;
        var stage = "rules.Create";
        try
        {
            logger.LogInformation(
                "Creating Outlook receive rule; ExistingRuleCount={ExistingRuleCount}, SenderTermCount={SenderTermCount}, SubjectTermCount={SubjectTermCount}, BodyTermCount={BodyTermCount}, BodyOrSubjectTermCount={BodyOrSubjectTermCount}, StopProcessing={StopProcessing}",
                originalCount, Count(request.SenderAddressContains), Count(request.SubjectContains), Count(request.BodyContains), Count(request.BodyOrSubjectContains), request.StopProcessingMoreRules);

            createdRule = rules.Create(request.RuleName, ReceiveRule);
            logger.LogInformation("Outlook rule stage completed: {RuleStage}", stage);

            stage = "subject/sender/body conditions";
            ConfigureReceiveRule((object)createdRule, destinationObject, request, logger);
            logger.LogInformation("Outlook rule stage completed: {RuleStage}", stage);

            stage = "rule.Enabled";
            createdRule.Enabled = true;
            if (!(bool)createdRule.Enabled) throw new InvalidOperationException("Outlook did not retain the enabled state for the new rule.");
            logger.LogInformation("Outlook rule stage completed: {RuleStage}", stage);

            stage = "rules.Save(false)";
            logger.LogInformation("Saving Outlook rule collection; RuleStage={RuleStage}", stage);
            rules.Save(false);
            logger.LogInformation("Outlook rule collection saved; RuleCount={RuleCount}", (int)rules.Count);
            return (object)createdRule;
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Outlook rule creation failed at stage {RuleStage}; starting exact-rule rollback", stage);
            Exception? rollbackFailure = null;
            if (createdRule is not null)
            {
                try
                {
                    RollBackCreatedRule(rulesObject, request.RuleName, originalCount);
                }
                catch (Exception rollbackException)
                {
                    rollbackFailure = rollbackException;
                    logger.LogCritical(rollbackException, "Outlook rule rollback failed after creation failed at stage {RuleStage}", stage);
                }
                finally
                {
                    ComReleaseHelper.FinalRelease(createdRule);
                }
            }

            if (rollbackFailure is not null)
            {
                throw new AggregateException("Outlook rule creation and exact-rule rollback both failed.", ex, rollbackFailure);
            }
            throw;
        }
    }

    internal static void ConfigureReceiveRule(object ruleObject, object destinationObject, CreateFolderRuleRequest request, ILogger logger)
    {
        dynamic rule = ruleObject;
        dynamic destination = destinationObject;
        dynamic? conditions = null;
        dynamic? actions = null;
        dynamic? condition = null;
        dynamic? move = null;
        dynamic? stop = null;
        try
        {
            conditions = rule.Conditions;
            if (HasValues(request.SenderAddressContains))
            {
                condition = conditions.SenderAddress;
                ConfigureStringArrayCondition(condition, "sender", request.SenderAddressContains!, true, logger);
                ComReleaseHelper.FinalRelease(condition);
                condition = null;
            }
            if (HasValues(request.SubjectContains))
            {
                condition = conditions.Subject;
                ConfigureStringArrayCondition(condition, "subject", request.SubjectContains!, false, logger);
                ComReleaseHelper.FinalRelease(condition);
                condition = null;
            }
            if (HasValues(request.BodyContains))
            {
                condition = conditions.Body;
                ConfigureStringArrayCondition(condition, "body", request.BodyContains!, false, logger);
                ComReleaseHelper.FinalRelease(condition);
                condition = null;
            }
            if (HasValues(request.BodyOrSubjectContains))
            {
                condition = conditions.BodyOrSubject;
                ConfigureStringArrayCondition(condition, "body_or_subject", request.BodyOrSubjectContains!, false, logger);
                ComReleaseHelper.FinalRelease(condition);
                condition = null;
            }

            actions = rule.Actions;
            try
            {
                move = actions.MoveToFolder;
                move.Enabled = true;
                SetComProperty((object)move, "Folder", destinationObject);
                if (!(bool)move.Enabled) throw new InvalidOperationException("Outlook did not retain the move-to-folder action's enabled state.");
                // Outlook can return null from MoveToFolder.Folder until Rules.Save has
                // materialized the action. A successful setter is the only reliable
                // pre-save check; the saved rule is verified independently afterwards.
                logger.LogInformation("Outlook rule element configured: move_to_folder; Enabled={Enabled}, DestinationAssigned={DestinationAssigned}", true, true);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Outlook rule element failed: move_to_folder");
                throw;
            }

            if (request.StopProcessingMoreRules)
            {
                try
                {
                    stop = actions.Stop;
                    stop.Enabled = true;
                    if (!(bool)stop.Enabled) throw new InvalidOperationException("Outlook did not retain the stop-processing action's enabled state.");
                    logger.LogInformation("Outlook rule element configured: stop_processing; Enabled={Enabled}", true);
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Outlook rule element failed: stop_processing");
                    throw;
                }
            }
            else
            {
                logger.LogInformation("Outlook rule element configured: stop_processing; Enabled={Enabled}", false);
            }
        }
        finally
        {
            ComReleaseHelper.FinalRelease(stop);
            ComReleaseHelper.FinalRelease(move);
            ComReleaseHelper.FinalRelease(condition);
            ComReleaseHelper.FinalRelease(actions);
            ComReleaseHelper.FinalRelease(conditions);
        }
    }

    private static void ConfigureStringArrayCondition(dynamic condition, string conditionName, IReadOnlyList<string> values, bool useAddressProperty, ILogger logger)
    {
        try
        {
            // Outlook exposes these properties as VARIANT. A managed string[] becomes
            // SAFEARRAY(BSTR), which Outlook accepts at assignment time but rejects when
            // Rules.Save validates the rule. object[] produces the VBA-compatible
            // SAFEARRAY(VARIANT), with each element marshalled as a BSTR.
            object[] variantArray = values.Select(static value => (object)value).ToArray();
            if (useAddressProperty) condition.Address = variantArray;
            else condition.Text = variantArray;
            condition.Enabled = true;

            var readBack = useAddressProperty ? (object)condition.Address : (object)condition.Text;
            VerifyStringArray(readBack, values.Count, conditionName);
            if (!(bool)condition.Enabled) throw new InvalidOperationException($"Outlook did not retain the {conditionName} condition's enabled state.");
            logger.LogInformation(
                "Outlook rule element configured: {ConditionName}; Enabled={Enabled}, TermCount={TermCount}, VariantArrayElementType={VariantArrayElementType}",
                conditionName, true, values.Count, typeof(object).Name);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Outlook rule element failed: {ConditionName}", conditionName);
            throw;
        }
    }

    private static void VerifyStringArray(object value, int expectedCount, string conditionName)
    {
        if (value is not IEnumerable sequence) throw new InvalidOperationException($"Outlook did not return an array for the {conditionName} condition.");
        var actualCount = sequence.Cast<object?>().Count();
        if (actualCount != expectedCount) throw new InvalidOperationException($"Outlook retained {actualCount} of {expectedCount} values for the {conditionName} condition.");
    }

    private static void SetComProperty(object target, string propertyName, object value)
    {
        _ = target.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, target, [value], System.Globalization.CultureInfo.InvariantCulture);
    }

    private void RollBackCreatedRule(object rulesObject, string ruleName, int originalCount)
    {
        dynamic rules = rulesObject;
        var removed = false;
        for (var index = 1; index <= (int)rules.Count; index++)
        {
            dynamic? rule = null;
            try
            {
                rule = rules[index];
                if (!string.Equals(Convert.ToString(rule.Name, System.Globalization.CultureInfo.InvariantCulture), ruleName, StringComparison.Ordinal)) continue;
                rules.Remove(index);
                removed = true;
                break;
            }
            finally
            {
                ComReleaseHelper.FinalRelease(rule);
            }
        }

        if (!removed) throw new InvalidOperationException("The newly created Outlook rule could not be found for rollback.");
        rules.Save(false);
        if ((int)rules.Count != originalCount) throw new InvalidOperationException("Outlook rule rollback did not restore the original rule count.");
        logger.LogWarning("Rolled back failed Outlook rule creation; RestoredRuleCount={RestoredRuleCount}", originalCount);
    }

    private static bool HasValues(IReadOnlyList<string>? values) => values is { Count: > 0 };
    private static int Count(IReadOnlyList<string>? values) => values?.Count ?? 0;
}
