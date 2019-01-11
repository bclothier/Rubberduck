﻿using System;
using System.Collections;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// A base class for exposing an enumerable collection through an index based accessor
    /// </summary>
    /// <typeparam name="TItem">the collection element type</typeparam>
    public abstract class IIndexedCollectionBase<TItem> : IEnumerable<TItem>
        where TItem : class
    {
        IEnumerator IEnumerable.GetEnumerator() => new IIndexedCollectionEnumerator<IIndexedCollectionBase<TItem>, TItem>(this);
        public IEnumerator<TItem> GetEnumerator() => new IIndexedCollectionEnumerator<IIndexedCollectionBase<TItem>, TItem>(this);

        public abstract int Count { get; }
        public abstract TItem GetItemByIndex(int index);
    }

    /// <summary>
    /// The enumerator implementation for IIndexedCollectionBase
    /// </summary>
    /// <typeparam name="TCollection">the IIndexedCollectionBase type</typeparam>
    /// <typeparam name="TItem">the collection element type</typeparam>
    public sealed class IIndexedCollectionEnumerator<TCollection, TItem> : IEnumerator<TItem>
        where TCollection : IIndexedCollectionBase<TItem>
        where TItem : class
    {
        private readonly TCollection _collection;
        private readonly int _collectionCount;
        private int _index = -1;
        private TItem _current;

        public IIndexedCollectionEnumerator(TCollection collection)
        {
            _collection = collection;
            _collectionCount = _collection.Count;
        }

        public void Dispose()
        {
            // nothing to do here.
        }

        TItem IEnumerator<TItem>.Current => _current;
        object IEnumerator.Current => _current;

        public void Reset() => _index = -1;

        public bool MoveNext()
        {
            _current = default(TItem);
            _index++;
            if (_index >= _collectionCount) return false;
            _current = _collection.GetItemByIndex(_index);
            return true;
        }
    }
}
