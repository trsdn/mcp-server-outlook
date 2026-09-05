using System.Diagnostics.CodeAnalysis;
using OutlookMcp.Core.Commands.OutlookInterop;
using OutlookMcp.Core.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMcp.Core.Commands.Oof;

/// <summary>
/// Read-only out-of-office (automatic replies) status backed by the <c>PR_OOF_STATE</c> store property.
/// </summary>
public class OofCommands : IOofCommands
{
    // PR_OOF_STATE / PidTagOutOfOfficeState: a boolean store property that is true when automatic
    // replies are switched on. It is the only out-of-office facet classic Outlook exposes through COM;
    // the reply bodies, the internal/external split and any scheduled window live in EWS/Graph only.
    private const string PrOofState = "http://schemas.microsoft.com/mapi/proptag/0x661D000B";

    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public OutlookOofStatusResult GetStatus()
    {
        return OutlookInteropRunner.Execute(
            "OutlookOofGetStatus",
            (application, session) =>
            {
                Outlook.Store? store = null;
                Outlook.PropertyAccessor? accessor = null;
                try
                {
                    store = session.DefaultStore;
                    string storeName = store?.DisplayName ?? string.Empty;
                    string storeType = ClassifyStoreType(store, out bool isExchange);

                    if (store == null || !isExchange)
                    {
                        // Not an error: out-of-office is an Exchange concept. A POP/IMAP or missing store
                        // simply does not support it, which is a domain answer, not a failure.
                        return new OutlookOofStatusResult
                        {
                            Success = true,
                            IsSupported = false,
                            StoreDisplayName = storeName,
                            ExchangeStoreType = storeType,
                            Note = "Out-of-office (automatic replies) applies only to an Exchange mailbox. "
                                + "The default store is not an Exchange store, so no out-of-office state is "
                                + "available."
                        };
                    }

                    accessor = store.PropertyAccessor;

                    // Let a genuine failure to read the property reach onException rather than masking it
                    // (Rule 1b). GetProperty returns a boxed System.Boolean for PR_OOF_STATE.
                    object raw = accessor.GetProperty(PrOofState);
                    bool enabled = Convert.ToBoolean(raw, System.Globalization.CultureInfo.InvariantCulture);

                    return new OutlookOofStatusResult
                    {
                        Success = true,
                        IsSupported = true,
                        IsOutOfOfficeEnabled = enabled,
                        StoreDisplayName = storeName,
                        ExchangeStoreType = storeType,
                        Note = "Read-only on/off state from PR_OOF_STATE. The reply text, the separate "
                            + "internal and external replies, and any scheduled start/end window are not "
                            + "exposed through Outlook COM (they require EWS or Microsoft Graph). In Cached "
                            + "Exchange mode this flag can lag the server until the next Send/Receive."
                    };
                }
                finally
                {
                    OutlookInteropRunner.ReleaseComObject(ref accessor);
                    OutlookInteropRunner.ReleaseComObject(ref store);
                }
            },
            ex => new OutlookOofStatusResult
            {
                Success = false,
                IsSupported = false,
                ErrorMessage = $"Failed to read the Outlook out-of-office status: {ex.Message}"
            });
    }

    /// <summary>
    /// Classifies the default store and reports whether it is an Exchange store. Reading
    /// <c>ExchangeStoreType</c> can throw on some stores, so a failure is treated as non-Exchange rather
    /// than failing the whole call - a best-effort classification, not error suppression of the operation.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    private static string ClassifyStoreType(Outlook.Store? store, out bool isExchange)
    {
        if (store == null)
        {
            isExchange = false;
            return "none";
        }

        Outlook.OlExchangeStoreType type;
        try
        {
            type = store.ExchangeStoreType;
        }
        catch
        {
            isExchange = false;
            return "unknown";
        }

        isExchange = type == Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox
            || type == Outlook.OlExchangeStoreType.olAdditionalExchangeMailbox
            || type == Outlook.OlExchangeStoreType.olExchangePublicFolder;

        return type.ToString();
    }
}
