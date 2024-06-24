#!/usr/bin/python
# coding=utf-8
import math

floor, ceil, sqrt = math.floor, math.ceil, math.sqrt
pow = math.pow
import operator
import pvhelper as pv


def simp_pearson(p1, p2):
    n = len(p1)
    assert n != 0
    sum1 = sum(p1)
    sum2 = sum(p2)
    m1 = float(sum1) / n
    m2 = float(sum2) / n
    p1mean = [(x - m1) for x in p1]
    p2mean = [(y - m2) for y in p2]
    numerator = sum(x * y for x, y in zip(p1mean, p2mean))
    denominator = math.sqrt(sum(x * x for x in p1mean) * sum(y * y for y in p2mean))
    return numerator / denominator if denominator else 0


def pearsR_vtime(vect):
    return simp_pearson(vect, list(range(len(vect))))


def trend_in_time(vect):
    return simp_pearson(vect, range(len(vect)))


def svar(X, method=0):
    """
    Computes the sample variance using various methods
    explained in the reference.

    method:
      0 - basic two-pass, the default.
      1 - single pass
      2 - compenstated
      3 - welford
    Reference: http://en.wikipedia.org/wiki/Algorithms_for_calculating_variance
    """
    if not (0 < method < 3):
        method = 0
    n = len(X)
    if method == 0:  # basic two-pass.
        xbar = sum(X) / float(n)
        return sum([(x - xbar) ** 2 for x in X]) / (n - 1.0)

    elif method == 1:  # single pass.
        S = 0.0
        SS = 0.0
        for x in X:
            S += x
            SS += x * x
        xbar = S / float(n)
        return (SS - n * xbar * xbar) / (n - 1.0)

    elif method == 2:  # compensated algorithm.
        xbar = mean(X)
        sum2 = 0
        sumc = 0
        for x in X:
            sum2 = sum2 + (x - mean) ^ 2
            sumc = sumc + (x - xbar)
        return (sum2 - sumc ^ 2 / n) / (n - 1)

    elif method == 3:  # Welford algorithm
        n = 0
        mean = 0
        M2 = 0

        for x in X:
            n = n + 1
            delta = x - mean
            mean = mean + delta / n
            M2 = M2 + delta * (x - mean)  # This expression uses the new value
        return M2 / (n - 1.0)


def sdev(X, method=1):
    return sqrt(svar(X, method))


def cov(X, Y):
    # covariance of X and Y vectors with same length.
    n = float(len(X))
    xbar = sum(X) / n
    ybar = sum(Y) / n
    return sum([(x - xbar) * (y - ybar) for x, y in zip(X, Y)]) / (n - 1)


Sxy = cov
Sx = sdev


def cor(X, Y):
    # correlation coefficient of X and Y.
    return cov(X, Y) / (Sx(X) * Sx(Y))


def average(values):
    return sum(values, 0.0) / float(len(values))


def percentile(N, percent, key=lambda x: x):
    """
    Find the percentile of a list of values.

    @parameter N - is a list of values. Note N MUST BE already sorted.
    @parameter percent - a float value from 0.0 to 1.0.
    @parameter key - optional key function to compute value from each element of N.

    @return - the percentile of the values
    """
    if not N:
        return None
    k = (len(N) - 1) * percent
    f = floor(k)
    c = ceil(k)
    if f == c:
        return key(N[int(k)])

    # changed to
    d0 = key(N[int(f)]) * (c - k)
    d1 = key(N[int(c)]) * (k - f)
    return d0 + d1


def get_distance(x, y):
    try:
        return math.sqrt((x - y) ** 2)
    except ValueError:
        return 0


def prc_diff_old(prev, today):
    prev, today = float(prev), float(today)
    try:
        return ((today - prev) / prev) * 100
    except ZeroDivisionError:
        return 0


def prc_diff(prev, today):
    try:
        return (today / float(prev) - 1) * 100
    except ZeroDivisionError:
        return 0


def get_prc_diffs(inlist, same_len=False):
    diffs = list(map(prc_diff, inlist[:-1], inlist[1:]))
    if same_len:
        diffs.insert(0, None)
    return diffs


def get_returns(dates, values):
    assert len(dates) == len(values)
    returns = get_prc_diffs(values)
    dates = dates[1:]  # cant calculate return for first date
    return dates, returns


def rank_simple(vector):
    return sorted(list(range(len(vector))), key=vector.__getitem__)


def rankdata(a):
    n = len(a)
    ivec = rank_simple(a)
    svec = [a[rank] for rank in ivec]
    sumranks = 0
    dupcount = 0
    newarray = [0] * n
    for i in range(n):
        sumranks += i
        dupcount += 1
        if i == n - 1 or svec[i] != svec[i + 1]:
            averank = sumranks / float(dupcount) + 1
            for j in range(i - dupcount + 1, i + 1):
                newarray[ivec[j]] = averank
            sumranks = 0
            dupcount = 0
    return newarray


def sort_ranked(inlist, cols_weights={}, adjfuncs={}, include_rank=False):
    if not inlist:
        raise Exception("Cant rank and sort empty list")
    if cols_weights and adjfuncs:
        if len(set(adjfuncs.keys()).difference(cols_weights)) != 0:
            raise Exception("adj function columns not in cols weights")
    if not cols_weights:
        return inlist

    def get_column(inlist, col):
        return list(map(operator.itemgetter(col), inlist))

    ranklist = []
    for idx, weight in cols_weights.items():
        data_col = get_column(inlist, idx)
        # ~ print "adjusting funct" print weight
        try:  # checking dictionary
            tmp = weight + 1.0
        except TypeError:
            # probably float then go standard hard coded to prevent errors
            funct = weight["funct"]
            weight = weight["weight"]
            data_col = list(map(funct, data_col))
        if idx in adjfuncs:
            data_col = list(map(adjfuncs[idx], data_col))
        ranklist.append([x * weight for x in rankdata(data_col)])
    sumlist = list(map(sum, list(zip(*ranklist))))
    if not include_rank:
        ranksAndInlist = list(zip(sumlist, inlist))
        ranksAndInlist.sort(reverse=True)
        sortedrank, sortedInlist = list(zip(*ranksAndInlist))
        return sortedInlist
    else:
        outlist = []
        for i in range(len(inlist)):
            outlist.append((sumlist[i],) + tuple(inlist[i]))
        outlist.sort(reverse=True)
        return outlist

    

def CAGR(start, end, years):
    assert_non_zero_years = 1 / years
    assert years > 2
    return (pow(end / float(start), 1 / float(years)) - 1) * 100


def get_sort_by_dict(rdict, flat_list, ticker_idx=0, verbose=True):
    rdict.get(None, False)
    missing_values = {}

    def mysort(row1, row2):
        if flat_list:  # single list not list of lists
            t1 = row1
            t2 = row2
        else:
            t1 = operator.itemgetter(ticker_idx)(row1)
            t2 = operator.itemgetter(ticker_idx)(row2)
        # ~ print t1, t2
        value1, value2 = rdict.get(t1), rdict.get(t2)
        if not value1:

            if t1 not in missing_values and verbose:
                print("""!could nof find values for sorting  value '%s'
                - check if list is flat
                """ % t1)
                # display message just once not on every compare
            missing_values[t1] = 1
            value1 = -99999

        if not value2:
            if t2 not in missing_values and verbose:

                print("""!could nof find values for sorting  value '%s'
                - check if list is flat
                """ % t2)
                # display message just once not on every compare
            missing_values[t2] = 1
            value2 = -99999

        assert_numbers = value1 + value2 + 1.0
        return cmp(value1, value2)

    return mysort


from collections import defaultdict
from operator import itemgetter
from math import fsum


def shimazaki_histogram(data, start, end):
    # http://toyoizumilab.brain.riken.jp/hideaki/res/histogram.html#Method
    _max = max(data)
    _min = min(data)
    results = []
    for N in range(start, end):
        width = float(_max - _min) / N

        hist = defaultdict(int)
        for x in data:
            i = int((x - _min) / width)
            if i >= N:  # Mimicking the behavior of matlab.histc(), and
                i = N - 1  # matplotlib.hist() and numpy.histogram().
            y = _min + width * i
            hist[y] += 1

        # Compute the mean and var.
        k = fsum(hist[x] for x in hist) / N
        v = fsum(hist[x] ** 2 for x in hist) / N - k**2

        C = (2 * k - v) / (width**2)

        results += [(hist, C, N, width)]

    optimal = min(results, key=itemgetter(1))

    if 0:  # if true, print bin-widths and C-values, the cost function.
        for hist, C, N, width in results:
            print("%f %f" % (width, C))

    return optimal


def test_hist_bins():
    data = [
        4.37,
        3.87,
        4.00,
        4.03,
        3.50,
        4.08,
        2.25,
        4.70,
        1.73,
        4.93,
        1.73,
        4.62,
        3.43,
        4.25,
        1.68,
        3.92,
        3.68,
        3.10,
        4.03,
        1.77,
        4.08,
        1.75,
        3.20,
        1.85,
        4.62,
        1.97,
        4.50,
        3.92,
        4.35,
        2.33,
        3.83,
        1.88,
        4.60,
        1.80,
        4.73,
        1.77,
        4.57,
        1.85,
        3.52,
        4.00,
        3.70,
        3.72,
        4.25,
        3.58,
        3.80,
        3.77,
        3.75,
        2.50,
        4.50,
        4.10,
        3.70,
        3.80,
        3.43,
        4.00,
        2.27,
        4.40,
        4.05,
        4.25,
        3.33,
        2.00,
        4.33,
        2.93,
        4.58,
        1.90,
        3.58,
        3.73,
        3.73,
        1.82,
        4.63,
        3.50,
        4.00,
        3.67,
        1.67,
        4.60,
        1.67,
        4.00,
        1.80,
        4.42,
        1.90,
        4.63,
        2.93,
        3.50,
        1.97,
        4.28,
        1.83,
        4.13,
        1.83,
        4.65,
        4.20,
        3.93,
        4.33,
        1.83,
        4.53,
        2.03,
        4.18,
        4.43,
        4.07,
        4.13,
        3.95,
        4.10,
        2.27,
        4.58,
        1.90,
        4.50,
        1.95,
        4.83,
        4.12,
    ]
    optimal = shimazaki_histogram(data, 4, 50)
    # Print the shimazaki optimal histogram.
    (hist, C, N, width) = optimal
    print("# C=%f, N=%d, width=%f" % (C, N, width))
    for x in sorted(hist):
        print("%f %d" % (x, hist[x]))


if __name__ == "__main__":
    print("zażółć gęślą jaźń")
    print(
        [
            1,
            2,
            3,
        ],
        test_hist_bins(),
    )

    print(CAGR(1, 3, 5))
    print(percentile(list(range(10)), 0.75))

    minval = -7.40
    max_val = 8.7

    ticks = 10.0

    vrange = get_distance(minval, max_val)

    step = vrange / ticks
    tick = minval
    while tick < max_val:
        tick += step
        position = get_distance(tick, minval) / float(vrange)
        print(tick, position)

    tmplist = [1, 1.1, 2, 2.05, 1, 0.9]

    print(tmplist)
    print(get_prc_diffs(tmplist))

    print(get_distance(-1, 4))
    a = list(range(100))
    b = list(range(100))
    c = (0,) * 100
    d = (0,) * 100
    from random import randint

    r = [randint(-10, 30) for x in a]
    r2 = [randint(-10, 30) for x in a]

    print(cor(a, b))
    print(simp_pearson(c, d))
    # ~ print cor(c,d)
    print(simp_pearson(c, r))
    print(sdev(r), sdev(c))
    print(simp_pearson(r, r2))
    print(percentile(a, 0.01))
    print(percentile(a, 0.98))
    print(percentile(a, 0.5))
