# Simulation-using-VBA-and-Macros-in-Excel
A VBA script for approximate optimization using simulation

Given an m-by-m matrix A = (aij) and a permu-tation π of {1, 2, . . . , m}, deﬁne the distance to be
d(π) = i = 1 to m Summation [aπ(i),π(i+1)]
where π(m + 1) ≡ π(1). 
The problem is to ﬁnd the permutation π which minimizes d(π). 
A simple approximate approach is to generate random permutations π(j), j = 1, 2, . . . , n and set
ˆ d(n) = minj=1,...,n d(π(j)). Use the 12-by-12 matrix A in the spreadsheet.
