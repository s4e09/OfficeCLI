// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// TrackingPropertyDictionary — wraps a property dictionary and records
// which keys the handler actually accessed (TryGetValue / ContainsKey /
// indexer / Remove). Used by the Add path to detect "user supplied
// --prop X=Y but the handler never read X" — which is the new
// definition of "unsupported property" under handler-as-truth.
//
// Architectural note: replaces the old SchemaHelpLoader.ValidateProperties
// pre-filter at CLI entry. Schema is no longer the runtime gate; the
// handler's actual consumption is. Aliases that the handler genuinely
// understands (whether or not the schema enumerates them) now flow
// through without warning. Real typos still produce a warning because
// the handler never reads them.
//
// Implementation note: we exploit Dictionary<TKey,TValue>'s use of
// IEqualityComparer<TKey>.Equals on every hash-based operation
// (TryGetValue, ContainsKey, indexer, Remove). The custom comparer
// records each lookup key. We seed the dictionary in the constructor
// before enabling recording so initial Add operations don't pollute
// the access set.
//
// Known leaks (acceptable for the typo-detection goal):
//  - foreach iteration: iterators don't go through the comparer, so a
//    handler that exhaustively foreaches the dict to find what it
//    wants won't mark anything as accessed. Mitigated by the new
//    GetEnumerator override below — when the static type is
//    TrackingPropertyDictionary, every yielded key is counted as
//    accessed. The override does NOT fire when the dict is upcast to
//    Dictionary<>; in that case foreach reads are silent. In practice
//    handlers iterate via for-loops over $"series{i}" probes, which
//    do go through TryGetValue, so this is rare.

using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace OfficeCli.Core;

internal sealed class TrackingPropertyDictionary : Dictionary<string, string>
{
    private readonly TrackingComparer _cmp;
    private readonly HashSet<string> _initialKeys;

    public TrackingPropertyDictionary(IDictionary<string, string> source)
        : base(new TrackingComparer(System.StringComparer.OrdinalIgnoreCase))
    {
        _cmp = (TrackingComparer)Comparer;
        foreach (var kv in source) base.Add(kv.Key, kv.Value);
        _initialKeys = new HashSet<string>(Keys, System.StringComparer.OrdinalIgnoreCase);
        _cmp.RecordingEnabled = true;
    }

    /// <summary>
    /// Keys the user supplied on the command line that the handler never
    /// touched via TryGetValue / ContainsKey / indexer / Remove. The
    /// caller surfaces these as <c>unsupported_property</c> warnings.
    /// </summary>
    public IReadOnlyCollection<string> UnusedKeys =>
        _initialKeys
            .Where(k => !_cmp.AccessedKeys.Contains(k))
            .ToList();

    /// <summary>Keys handler accessed (subset of input ∪ keys it added).</summary>
    public IReadOnlyCollection<string> AccessedKeys => _cmp.AccessedKeys;

    public new IEnumerator<KeyValuePair<string, string>> GetEnumerator()
    {
        foreach (var kv in (IDictionary<string, string>)this)
        {
            _cmp.AccessedKeys.Add(kv.Key);
            yield return kv;
        }
    }

    private sealed class TrackingComparer : IEqualityComparer<string>
    {
        private readonly IEqualityComparer<string> _inner;
        public bool RecordingEnabled;
        public readonly HashSet<string> AccessedKeys =
            new(System.StringComparer.OrdinalIgnoreCase);

        public TrackingComparer(IEqualityComparer<string> inner) => _inner = inner;

        public bool Equals(string? x, string? y)
        {
            if (RecordingEnabled)
            {
                // Dictionary<,> calls Equals(lookup_key, stored_key). Both
                // refer to the same logical key (case-insensitive); record
                // the canonical (stored) form so we don't double-count
                // case variants.
                if (y != null) AccessedKeys.Add(y);
            }
            return _inner.Equals(x, y);
        }

        public int GetHashCode(string obj) => _inner.GetHashCode(obj);
    }
}
