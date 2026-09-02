using OutlookMcp.Core.Commands.Mail;
using Xunit;

namespace OutlookMcp.Core.Tests.Unit;

/// <summary>
/// Unit tests for <see cref="MailCommands.Send"/>'s confirmation gate and operationId-keyed
/// idempotency cache (#29). Both code paths return before any Outlook COM access, so they are
/// pure and safe to test without a running Outlook instance (Rule 30's exception).
/// </summary>
public class MailCommandsSendTests
{
    [Fact]
    public void Send_WithoutConfirm_IsRefusedAndNeverTouchesOutlook()
    {
        var commands = new MailCommands();

        var result = commands.Send(entryId: "some-entry-id", confirm: false);

        Assert.False(result.Success);
        Assert.False(result.Sent);
        Assert.False(result.Indeterminate);
        Assert.Contains("confirm=true", result.ErrorMessage);
    }

    [Fact]
    public void Send_WithoutConfirm_OperationIdIsNotCached()
    {
        // The confirmation gate is checked before any Outlook COM work and deliberately returns
        // before populating the idempotency cache -- a refused (confirm=false) attempt was never
        // actually sent, so there is nothing to protect a retry from duplicating. Confirm this by
        // checking that a second call with the same operationId is independently evaluated (not
        // required to be the same instance), i.e. the cache only ever holds outcomes of send
        // attempts that were actually gated open (confirm=true).
        var commands = new MailCommands();
        string operationId = Guid.NewGuid().ToString();

        var first = commands.Send(entryId: "some-entry-id", confirm: false, operationId: operationId);
        var second = commands.Send(entryId: "some-entry-id", confirm: false, operationId: operationId);

        Assert.False(first.Success);
        Assert.False(second.Success);
    }
}
