import numpy as np

cox_sprague_table = [
    [100, 99, 98, 97, 96, 94, 92, 90, 87, 84, 80, 76, 72, 66, 60, 52, 43, 31, 10],
    [94, 93, 92, 91, 90, 88, 86, 84, 81, 78, 74, 70, 66, 60, 54, 46, 37, 25, 7],
    [90, 89, 88, 87, 86, 84, 82, 80, 77, 74, 70, 66, 62, 56, 50, 42, 33, 21, 5],
    [86, 85, 84, 83, 82, 80, 78, 76, 73, 70, 66, 62, 58, 52, 46, 38, 29, 17],
    [83, 82, 81, 80, 79, 77, 75, 73, 70, 67, 63, 59, 55, 49, 43, 35, 26],
    [80, 79, 78, 77, 76, 74, 72, 70, 67, 64, 60, 56, 52, 46, 40, 32],
    [78, 77, 76, 75, 74, 72, 70, 68, 65, 62, 58, 54, 50, 44, 38],
    [76, 75, 74, 73, 72, 70, 68, 66, 63, 60, 56, 52, 48, 42],
    [74, 73, 72, 71, 70, 68, 66, 64, 61, 58, 54, 50, 46],
    [72, 71, 70, 69, 68, 66, 64, 62, 59, 56, 52, 48],
    [70, 69, 68, 67, 66, 64, 62, 60, 57, 54, 50],
    [68, 67, 66, 65, 64, 62, 60, 58, 55, 52],
    [66, 65, 64, 63, 62, 60, 58, 56, 53],
    [64, 63, 62, 61, 60, 58, 56, 54],
    [62, 61, 60, 59, 58, 56, 54],
    [60, 59, 58, 57, 56, 54],
    [58, 57, 56, 55, 54],
    [56, 55, 54, 53],
    [54, 53, 52],
    [52, 51]
]

"""
Since the table is expected to be a 100x20 array, we would need to extend
the above matrix with the remaining data. Assuming we have the rest of the data,
the full matrix would be 100 rows, with each row having up to 20 columns.
Fill in the missing data as required to complete the table.

https://www.shieldsfleetone.org/Cox-Sprague_Scoring.htm

Cox-Sprague Table
Place                 Number of Starters
         2  3  4  5  6  7  8  9 10 11 12 13 14 15 16 17 18 19 20

 1      10 31 43 52 60 66 72 76 80 84 87 90 92 94 96 97 98 99 100
 2       7 25 37 46 54 60 66 70 74 78 81 84 86 88 90 91 92 93 94
 3       5 21 33 42 50 56 62 66 70 74 77 80 82 84 86 87 88 89 90
 4       . 17 29 38 46 52 58 62 66 70 73 76 78 80 82 83 84 85 86
 5       .  . 26 35 43 49 55 59 63 67 70 73 75 77 79 80 81 82 83
 6       .  .  . 32 40 46 52 56 60 64 67 70 72 74 76 77 78 79 80
 7       .  .  .  . 38 44 50 54 58 62 65 68 70 72 74 75 76 77 78
 8       .  .  .  .  . 42 48 52 56 60 63 66 68 70 72 73 74 75 76
 9       .  .  .  .  .  . 46 50 54 58 61 64 66 68 70 71 72 73 74
10       .  .  .  .  .  .  . 48 52 56 59 62 64 66 68 69 70 71 72
11       .  .  .  .  .  .  .  . 50 54 57 60 62 64 66 67 68 69 70
12       .  .  .  .  .  .  .  .  . 52 55 58 60 62 64 65 66 67 68
13       .  .  .  .  .  .  .  .  .  . 53 56 58 60 62 63 64 65 66
14       .  .  .  .  .  .  .  .  .  .  . 55 57 59 61 62 63 64 65
15       .  .  .  .  .  .  .  .  .  .  .  . 56 58 60 61 62 63 64
16       .  .  .  .  .  .  .  .  .  .  .  .  . 57 59 60 61 62 63
17       .  .  .  .  .  .  .  .  .  .  .  .  .  . 58 59 60 61 62
18       .  .  .  .  .  .  .  .  .  .  .  .  .  .  . 58 59 60 61
19       .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . 58 59 60
20       .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . 58 59
21       .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  . 58
"""
def csScore(resultsVector, nStartersVector, CoxSpragueTable, nFinishersVector=None, nDiscards=0):
    nx = 0  # counter of number of races started by the boat being scored
    xTot = 0  # running C-S score total
    pTot = 0  # running C-S "perfect" total
    xv = []  # vector C-S scores in races started (for computing discards)
    pv = []  # vector of perfect C-S scores (for computing discards)

    for i, ns in enumerate(nStartersVector):
        if ns > 0:
            c = min(ns, 20)  # Adjust column index to the Cox-Sprague table limit
            result = resultsVector[i]
            if type(result) == int:  # Handling numeric results
                r = result
            else:  # Handling special cases as strings
                r = ns + 1  # Default scoring for DSQ/DNF/etc.
                if result in ['DNS', 'DNC']:  # Ignore DNS/DNC
                    continue
            # Adjust score based on the Cox-Sprague table
            z = CoxSpragueTable[r-1][c-1]  # Adjusted for Python's 0-based indexing
            nx += 1
            xv.append(z)
            xTot += z
            pv.append(CoxSpragueTable[0][c-1])  # Score for winning
            pTot += CoxSpragueTable[0][c-1]

    # Calculate Cox-Sprague percentage of perfection before discards
    cs = xTot / pTot if pTot > 0 else 0

    # Handle discards if necessary
    for _ in range(min(nDiscards, nx - 1)):
        discard_improvements = [(xTot - x) / (pTot - p) if pTot - p > 0 else 0 for x, p in zip(xv, pv)]
        im = np.argmax(discard_improvements)  # Index of the worst race to discard
        if discard_improvements[im] >= cs:
            xTot -= xv[im]
            pTot -= pv[im]
            xv[im] = 0  # Mark as discarded

    return cs

def csTable1(r, c, modified=True):
    # Simplified version, needs to be filled with actual Cox-Sprague table data
    # For demonstration, a simplified table can be created with dummy values
    if c > 20:
        c = 20
    # Example of creating a simplified Cox-Sprague table
    table = np.full((100, 20), fill_value=0.0)  # Dummy table, replace with actual data
    # Calculate and return value from the table
    # Add logic here to match the VBA function's behavior
    return table[r-1][c-1]

# Example usage:
resultsVector = [1, 2, 'DSQ', 4]  # Example data
nStartersVector = [5, 10, 15, 20]
cox_sprague_table
nFinishersVector = [4, 9, 14, 19]  # Example data

score = csScore(resultsVector, nStartersVector, cox_sprague_table, nFinishersVector, nDiscards=1)
print(f"Cox-Sprague Score: {score}")
