import gamspy as gp
from numpy.random import uniform

m = gp.Container()
i = gp.Set(m)
j = gp.Set(m)
a = gp.Parameter(m, domain=[i, j])
b = gp.Parameter(m, domain=[i])
x = gp.Variable(m, domain=[i, j])
e = gp.Equation(m, domain=[i])
e[i] = gp.Sum(j, a[i, j] * x[i, j]) >= b[i]

data = uniform(0, 1, (500, 1000))
data[data > 0.01] = 0
i.setRecords(range(500))
j.setRecords(range(1000))
a.setRecords(data)
b.setRecords(uniform(0, 1, 500))