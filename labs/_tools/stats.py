# /// script
# requires-python = ">=3.10"
# dependencies = []
# ///
"""Exact figures quoted on the slides and in the practice bank: Wilson intervals of the assistant
error rates, the stop-rule probabilities and the worked examples. Standard library only.

    uv run labs/_tools/stats.py
"""
from math import comb, sqrt


def tail(n, k, p):
    """P(X >= k) for X ~ Binomial(n, p), exact."""
    return sum(comb(n, i) * p ** i * (1 - p) ** (n - i) for i in range(k, n + 1))


def wilson(x, n, z=1.959963984540054):
    p = x / n
    den = 1 + z * z / n
    centre = (p + z * z / (2 * n)) / den
    half = z * sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / den
    return centre - half, centre + half


print("Passing rule of the supervised check: 9 of 12 (70% of 12 is 8.4)")
print(f"  P(>=9 of 12 by guessing among three labels) = {tail(12, 9, 1/3):.6f}")
print(f"  P(>=51 of 60 by guessing among three labels) = {tail(60, 51, 1/3):.3e}")
print()
print("Assistant error rates, Wilson 95% intervals")
for label, x, n in [("batch 1 all", 9, 40), ("batch 1 policies", 1, 14), ("batch 1 reports", 6, 12),
                    ("batch 1 pages", 2, 14), ("batch 2 all", 11, 60), ("batch 2 policies", 1, 20),
                    ("batch 2 reports", 7, 20), ("batch 2 pages", 3, 20), ("both policies", 2, 34),
                    ("both reports", 13, 32), ("both pages", 5, 34)]:
    lo, hi = wilson(x, n)
    print(f"  {label:18s} {x:2d}/{n:2d} = {100*x/n:5.1f}%  [{100*lo:4.1f}%, {100*hi:4.1f}%]")
print()
print("Re-measuring rule: 20 signed-off mappings per cycle, stop the type at 3 or more wrong")
for p in (0.05, 0.06, 0.10, 0.15, 0.25, 0.40):
    print(f"  true error rate {p:.2f}: P(stop) = {tail(20, 3, p):.3f}")
print()
print("Slide parallel (Deanship of Graduate Studies): 15 signed-off mappings, stop at 2 or more wrong")
for p in (0.05, 0.10, 0.25):
    print(f"  true error rate {p:.2f}: P(stop) = {tail(15, 2, p):.3f}")
print("Slide parallel: forms 1/20 and guides 6/20, Wilson 95% intervals")
for label, x, n in [("forms", 1, 20), ("guides", 6, 20), ("housing 3/12 high", 3, 12), ("housing 1/8 low", 1, 8)]:
    lo, hi = wilson(x, n)
    print(f"  {label:18s} {x:2d}/{n:2d} = {100*x/n:5.1f}%  [{100*lo:4.1f}%, {100*hi:4.1f}%]")
print("Bank: library 21/25 agreement = 84%; target 85% of 25 = 21.25, so 22 items")
print(f"Slide median example: values 3,5,6,9,14,40,90 -> median 9, mean {sum([3,5,6,9,14,40,90])/7:.1f}")
