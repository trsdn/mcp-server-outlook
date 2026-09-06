namespace OutlookMcp.Core.Commands.Mail;

/// <summary>
/// Tracks the tied band at the frontier of a listing scan, so the cursor minted at the end of a page
/// carries <b>every</b> id already served at that instant rather than only the ones this page
/// happened to return (#135).
///
/// <para>
/// <b>The bug this exists to remove.</b> Paging is a keyset walk over received time, and received
/// times are not unique. <see cref="MailPageCursor.Includes"/> deliberately re-scans the boundary
/// instant inclusively and excludes what has already been served <i>by identity</i>, which is only
/// correct if the cursor lists every id served at that instant across every page that shared it -
/// exactly as <see cref="MailPageCursor.Encode"/> requires.
/// </para>
///
/// <para>
/// Both listing paths used to accumulate those ids in a list local to one call, cleared the moment
/// the frontier timestamp was first seen. On a resumed page that clearing happens immediately, so
/// each cursor carried only its own page's ids and forgot the previous page's. With three messages
/// A, B and C sharing instant T and a page size of one, the walk goes A, then B (cursor
/// <c>T:[B]</c>, having forgotten A), then A again, then B again - forever. C is never reached,
/// <c>hasMore</c> never goes false, and every individual response looks perfectly well-formed and
/// reports success. A caller looping until the cursor runs out does not terminate.
/// </para>
///
/// <para>
/// The fix is to carry the incoming cursor's ids forward while the frontier is still that same
/// instant, and to forget them only on genuinely advancing to an older one. Kept here, as one
/// deliberately COM-free type, because the alternative was a second copy of this reasoning in the
/// other listing path - and it is a second copy of it that produced the bug.
/// </para>
/// </summary>
internal sealed class MailPageBoundary
{
    private readonly DateTimeOffset? _inheritedInstant;
    private readonly IReadOnlyList<string> _inheritedIds;
    private readonly List<string> _ids = [];

    /// <param name="cursor">
    /// The cursor this page is resuming from, or <see langword="null"/> for a first page - in which
    /// case there is nothing to inherit, since no id has been served yet.
    /// </param>
    public MailPageBoundary(MailPageCursor? cursor)
    {
        _inheritedInstant = cursor?.LastReceived;
        _inheritedIds = cursor?.SeenAtBoundary ?? [];
    }

    /// <summary>The frontier instant, or <see langword="null"/> if nothing was observed.</summary>
    public DateTimeOffset? Instant { get; private set; }

    /// <summary>
    /// Every id served at <see cref="Instant"/>, across this page and any earlier page that shared
    /// the instant. This is what the next cursor must carry.
    /// </summary>
    public IReadOnlyList<string> Ids => _ids;

    /// <summary>
    /// Records a message that this page is about to return, in scan order (received time
    /// descending).
    /// </summary>
    /// <param name="receivedUtc">The message's received time, normalised to UTC.</param>
    /// <param name="entryId">Its entry id.</param>
    public void Observe(DateTimeOffset receivedUtc, string entryId)
    {
        if (Instant != receivedUtc)
        {
            Instant = receivedUtc;
            _ids.Clear();

            // The previous page stopped at this same instant and already served some of the band.
            // Those ids have to stay in the cursor: the resume filter excludes by identity, so an id
            // dropped here is served a second time, and the walk oscillates instead of advancing.
            if (_inheritedInstant == receivedUtc)
            {
                _ids.AddRange(_inheritedIds);
            }
        }

        // Guarded rather than assumed. A duplicate would not corrupt the walk - the resume filter
        // treats the list as a set - but it would grow every cursor in a long tied band without
        // bound.
        if (!_ids.Contains(entryId, StringComparer.Ordinal))
        {
            _ids.Add(entryId);
        }
    }
}
