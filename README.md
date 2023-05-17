Sub specanalysis()

'İnput
'Newmark Method
'k is taken 1
'Avarage acceleration method was udes in the calculation of a1, a2, a3 and kes values

Set syf1 = Worksheets("Acceleration")
Set syf4 = Worksheets("Initial Calculation")
Set syf5 = Worksheets("For each time step")

ns = syf4.Cells(1, 2)
nt = syf4.Cells(2, 2)
gama = syf4.Cells(3, 2)
beta = syf4.Cells(4, 2)
ntt = syf4.Cells(5, 2)

Dim ivmeveri() As Double
ReDim ivmeveri(1 To 12600, 1 To 10)


For j = 1 To 10
For i = 1 To 12600

ivmeveri(i, j) = syf1.Cells(i + 2, j)

Next i
Next j

Dim n() As Double, m() As Double, w() As Double, t1() As Double, ccr() As Double, c() As Double, kes() As Double, a1() As Double, a2() As Double, a3() As Double
ReDim n(0 To nt), m(0 To nt), w(0 To nt), t1(0 To nt), ccr(0 To nt), c(0 To nt), kes(0 To nt), a1(0 To nt), a2(0 To nt), a3(0 To nt)
Dim u() As Double, uu() As Double, uuu() As Double, usonucu() As Double, vsonucu() As Double, asonucu() As Double
ReDim u(0 To 12601, 1 To 10, 1 To nt), uu(0 To 12601, 1 To 10, 1 To nt), uuu(0 To 12601, 1 To 10, 1 To nt), usonucu(0 To nt, 1 To 10), vsonucu(0 To nt, 1 To 10), asonucu(0 To nt, 1 To 10)


For i = 1 To nt

n(i) = i * ns
m(i) = (n(i) / (2 * 3.14)) ^ 2
w(i) = (1 / m(i)) ^ (0.5)
t1(i) = 2 * 3.14 / w(i)
ccr(i) = 2 * m(i) * w(i)
c(i) = 0.05 * ccr(i) '%5 damp ratio
a1(i) = 40000 * m(i) + 200 * c(i)
a2(i) = 400 * m(i) + c(i)
a3(i) = m(i)
kes(i) = a1(i) + 1
byk2 = 0
byk3 = 0
byk4 = 0
byk5 = 0
byk6 = 0
byk7 = 0
byk8 = 0
byk9 = 0
byk10 = 0

For j = 2 To 12600

Piartı12 = -m(i) * ivmeveri(j, 2)
ppiartı12 = Piartı12 + a1(i) * u(j, 2, i) + a2(i) * uu(j, 2, i) + a3(i) * uuu(j, 2, i)
u(j + 1, 2, i) = (ppiartı12 / kes(i))
uu(j + 1, 2, i) = 200 * (u(j + 1, 2, i) - u(j, 2, i)) - (uu(j, 2, i))
uuu(j + 1, 2, i) = 40000 * (u(j + 1, 2, i) - u(j, 2, i)) - 400 * uu(j, 2, i) - 1 * uuu(j, 2, i)

dgr2 = Abs(u(j, 2, i))
If dgr2 > byk2 Then byk2 = dgr2


Piartı13 = -m(i) * ivmeveri(j, 3)
ppiartı13 = Piartı13 + a1(i) * u(j, 3, i) + a2(i) * uu(j, 3, i) + a3(i) * uuu(j, 3, i)
u(j + 1, 3, i) = (ppiartı13 / kes(i))
uu(j + 1, 3, i) = 200 * (u(j + 1, 3, i) - u(j, 3, i)) - (uu(j, 3, i))
uuu(j + 1, 3, i) = 40000 * (u(j + 1, 3, i) - u(j, 3, i)) - 400 * uu(j, 3, i) - 1 * uuu(j, 3, i)

dgr3 = Abs(u(j, 3, i))
If dgr3 > byk3 Then byk3 = dgr3


Piartı14 = -m(i) * ivmeveri(j, 4)
ppiartı14 = Piartı14 + a1(i) * u(j, 4, i) + a2(i) * uu(j, 4, i) + a3(i) * uuu(j, 4, i)
u(j + 1, 4, i) = (ppiartı14 / kes(i))
uu(j + 1, 4, i) = 200 * (u(j + 1, 4, i) - u(j, 4, i)) - (uu(j, 4, i))
uuu(j + 1, 4, i) = 40000 * (u(j + 1, 4, i) - u(j, 4, i)) - 400 * uu(j, 4, i) - 1 * uuu(j, 4, i)

dgr4 = Abs(u(j, 4, i))
If dgr4 > byk4 Then byk4 = dgr4


Piartı15 = -m(i) * ivmeveri(j, 5)
ppiartı15 = Piartı15 + a1(i) * u(j, 5, i) + a2(i) * uu(j, 5, i) + a3(i) * uuu(j, 5, i)
u(j + 1, 5, i) = (ppiartı15 / kes(i))
uu(j + 1, 5, i) = 200 * (u(j + 1, 5, i) - u(j, 5, i)) - (uu(j, 5, i))
uuu(j + 1, 5, i) = 40000 * (u(j + 1, 5, i) - u(j, 5, i)) - 400 * uu(j, 5, i) - 1 * uuu(j, 5, i)

dgr5 = Abs(u(j, 5, i))
If dgr5 > byk5 Then byk5 = dgr5


Piartı16 = -m(i) * ivmeveri(j, 6)
ppiartı16 = Piartı16 + a1(i) * u(j, 6, i) + a2(i) * uu(j, 6, i) + a3(i) * uuu(j, 6, i)
u(j + 1, 6, i) = (ppiartı16 / kes(i))
uu(j + 1, 6, i) = 200 * (u(j + 1, 6, i) - u(j, 6, i)) - (uu(j, 6, i))
uuu(j + 1, 6, i) = 40000 * (u(j + 1, 6, i) - u(j, 6, i)) - 400 * uu(j, 6, i) - 1 * uuu(j, 6, i)

dgr6 = Abs(u(j, 6, i))
If dgr6 > byk6 Then byk6 = dgr6


Piartı17 = -m(i) * ivmeveri(j, 7)
ppiartı17 = Piartı17 + a1(i) * u(j, 7, i) + a2(i) * uu(j, 7, i) + a3(i) * uuu(j, 7, i)
u(j + 1, 7, i) = (ppiartı17 / kes(i))
uu(j + 1, 7, i) = 200 * (u(j + 1, 7, i) - u(j, 7, i)) - (uu(j, 7, i))
uuu(j + 1, 7, i) = 40000 * (u(j + 1, 7, i) - u(j, 7, i)) - 400 * uu(j, 7, i) - 1 * uuu(j, 7, i)

dgr7 = Abs(u(j, 7, i))
If dgr7 > byk7 Then byk7 = dgr7



Piartı18 = -m(i) * ivmeveri(j, 8)
ppiartı18 = Piartı18 + a1(i) * u(j, 8, i) + a2(i) * uu(j, 8, i) + a3(i) * uuu(j, 8, i)
u(j + 1, 8, i) = (ppiartı18 / kes(i))
uu(j + 1, 8, i) = 200 * (u(j + 1, 8, i) - u(j, 8, i)) - (uu(j, 8, i))
uuu(j + 1, 8, i) = 40000 * (u(j + 1, 8, i) - u(j, 8, i)) - 400 * uu(j, 8, i) - 1 * uuu(j, 8, i)

dgr8 = Abs(u(j, 8, i))
If dgr8 > byk8 Then byk8 = dgr8


Piartı19 = -m(i) * ivmeveri(j, 9)
ppiartı19 = Piartı19 + a1(i) * u(j, 9, i) + a2(i) * uu(j, 9, i) + a3(i) * uuu(j, 9, i)
u(j + 1, 9, i) = (ppiartı19 / kes(i))
uu(j + 1, 9, i) = 200 * (u(j + 1, 9, i) - u(j, 9, i)) - (uu(j, 9, i))
uuu(j + 1, 9, i) = 40000 * (u(j + 1, 9, i) - u(j, 9, i)) - 400 * uu(j, 9, i) - 1 * uuu(j, 9, i)

dgr9 = Abs(u(j, 9, i))
If dgr9 > byk9 Then byk9 = dgr9


Piartı110 = -m(i) * ivmeveri(j, 10)
ppiartı110 = Piartı110 + a1(i) * u(j, 10, i) + a2(i) * uu(j, 10, i) + a3(i) * uuu(j, 10, i)
u(j + 1, 10, i) = (ppiartı110 / kes(i))
uu(j + 1, 10, i) = 200 * (u(j + 1, 10, i) - u(j, 10, i)) - (uu(j, 10, i))
uuu(j + 1, 10, i) = 40000 * (u(j + 1, 10, i) - u(j, 10, i)) - 400 * uu(j, 10, i) - 1 * uuu(j, 10, i)

dgr10 = Abs(u(j, 10, i))
If dgr10 > byk10 Then byk10 = dgr10


usonucu(i, 2) = byk2
usonucu(i, 3) = byk3
usonucu(i, 4) = byk4
usonucu(i, 5) = byk5
usonucu(i, 6) = byk6
usonucu(i, 7) = byk7
usonucu(i, 8) = byk8
usonucu(i, 9) = byk9
usonucu(i, 10) = byk10

Next j
Next i


For i = 1 To nt
For j = 2 To 10

vsonucu(i, j) = ((2 * 3.14) / t1(i)) * usonucu(i, j)
asonucu(i, j) = ((2 * 3.14) / t1(i)) ^ 2 * usonucu(i, j)

Next j
Next i


For i = 1 To nt

syf5.Cells(i + 3, 2) = usonucu(i, 2)
syf5.Cells(i + 3, 3) = vsonucu(i, 2)
syf5.Cells(i + 3, 4) = asonucu(i, 2)

syf5.Cells(i + 3, 5) = usonucu(i, 3)
syf5.Cells(i + 3, 6) = vsonucu(i, 3)
syf5.Cells(i + 3, 7) = asonucu(i, 3)

syf5.Cells(i + 3, 8) = usonucu(i, 4)
syf5.Cells(i + 3, 9) = vsonucu(i, 4)
syf5.Cells(i + 3, 10) = asonucu(i, 4)

syf5.Cells(i + 3, 11) = usonucu(i, 5)
syf5.Cells(i + 3, 12) = vsonucu(i, 5)
syf5.Cells(i + 3, 13) = asonucu(i, 5)

syf5.Cells(i + 3, 14) = usonucu(i, 6)
syf5.Cells(i + 3, 15) = vsonucu(i, 6)
syf5.Cells(i + 3, 16) = asonucu(i, 6)

syf5.Cells(i + 3, 17) = usonucu(i, 7)
syf5.Cells(i + 3, 18) = vsonucu(i, 7)
syf5.Cells(i + 3, 19) = asonucu(i, 7)

syf5.Cells(i + 3, 20) = usonucu(i, 8)
syf5.Cells(i + 3, 21) = vsonucu(i, 8)
syf5.Cells(i + 3, 22) = asonucu(i, 8)

syf5.Cells(i + 3, 23) = usonucu(i, 9)
syf5.Cells(i + 3, 24) = vsonucu(i, 9)
syf5.Cells(i + 3, 25) = asonucu(i, 9)

syf5.Cells(i + 3, 26) = usonucu(i, 10)
syf5.Cells(i + 3, 27) = vsonucu(i, 10)
syf5.Cells(i + 3, 28) = asonucu(i, 10)

Next i


End Sub









<!---
masici-ce/masici-ce is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
