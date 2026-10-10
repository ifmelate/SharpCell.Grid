using System;
using System.Collections.Generic;

namespace SharpCell.Grid;

/// <summary>
/// Sizes and offsets along one axis of a sheet: the rows or the columns. Every item has the
/// default size unless set otherwise; only the changed ones are stored, so a sheet of a million
/// rows costs nothing until rows get their own heights.
/// </summary>
internal sealed class AxisLayout
{
    private readonly int count;
    private readonly string name;
    private readonly SortedDictionary<int, double> sizes = new();
    private double defaultSize;

    // Built from sizes when an offset is asked for after a change: the changed indices in order and,
    // for each, the sum of (size - default) over it and all changed indices before it.
    private int[] changed = [];
    private double[] deltaThrough = [];
    private bool stale;

    public AxisLayout(int count, double defaultSize, string name)
    {
        this.count = count;
        this.name = name;
        DefaultSize = defaultSize;
    }

    public int Count => count;

    public double DefaultSize
    {
        get => defaultSize;
        set
        {
            if (!(value > 0) || double.IsInfinity(value))
                throw new ArgumentOutOfRangeException(nameof(value), value, $"The default {name} size must be positive.");
            defaultSize = value;
            stale = true;
        }
    }

    /// <summary>The offset where everything ends: the sum of all sizes.</summary>
    public double Total => OffsetOf(count) + SizeOf(count);

    public double SizeOf(int index)
    {
        CheckIndex(index);
        return sizes.TryGetValue(index, out var size) ? size : defaultSize;
    }

    public void SetSize(int index, double size)
    {
        CheckIndex(index);
        if (!(size >= 0) || double.IsInfinity(size))
            throw new ArgumentOutOfRangeException(nameof(size), size, $"A {name} size must be zero or more.");
        if (size == defaultSize)
            sizes.Remove(index);
        else
            sizes[index] = size;
        stale = true;
    }

    /// <summary>Where the item starts: 0 for the first one.</summary>
    public double OffsetOf(int index)
    {
        CheckIndex(index);
        Refresh();
        var before = CountChangedBefore(index);
        return (index - 1) * defaultSize + (before == 0 ? 0 : deltaThrough[before - 1]);
    }

    /// <summary>
    /// The item whose span holds the offset; items of size zero never do. Offsets before the start
    /// give the first item, offsets past the end the last.
    /// </summary>
    public int IndexAt(double offset)
    {
        if (double.IsNaN(offset))
            throw new ArgumentOutOfRangeException(nameof(offset), offset, "The offset is not a number.");

        // The last index that starts at or before the offset; of items starting at the same place,
        // only the last can have a size, so it is the one hit.
        int low = 1, high = count;
        while (low < high)
        {
            var middle = low + (high - low + 1) / 2;
            if (OffsetOf(middle) <= offset)
                low = middle;
            else
                high = middle - 1;
        }
        return low;
    }

    private int CountChangedBefore(int index)
    {
        var position = Array.BinarySearch(changed, index);
        return position >= 0 ? position : ~position;
    }

    private void Refresh()
    {
        if (!stale)
            return;
        changed = new int[sizes.Count];
        deltaThrough = new double[sizes.Count];
        double sum = 0;
        var i = 0;
        foreach (var (index, size) in sizes)
        {
            sum += size - defaultSize;
            changed[i] = index;
            deltaThrough[i] = sum;
            i++;
        }
        stale = false;
    }

    private void CheckIndex(int index)
    {
        if (index < 1 || index > count)
            throw new ArgumentOutOfRangeException(nameof(index), index, $"The {name} is outside the sheet.");
    }
}
