import openpyxl
import numpy as np
import time

def importData(sheet):
    return ([[cell.value for cell in row] for row in sheet])


wb = openpyxl.load_workbook("Data.xlsx")

def importFuzzyData(data):
    fuzzyData = wb['FuzzyData2']
    base = list(fuzzyData.values)
    return base



def combination(k, n):
    if k == 0 or k == n:
        return 1
    if k == 1:
        return n
    return combination(k - 1, n - 1) + combination(k, n - 1)


def caculateA(base):
    colum = len(base[0])
    row = len(base)
    A = np.zeros((row, combination(4, colum - 1)))

    for r1 in range(row):
        k = [0] * combination(4, colum - 1)
        temp = 0
        for a in range(0, colum - 4):
            for b in range(a + 1, colum - 3):
                for c in range(b + 1, colum - 2):
                    for d in range(c + 1, colum - 1):
                        for r2 in range(row):
                            if base[r1][a] == base[r2][a] and base[r1][b] == base[r2][b] and base[r1][c] == base[r2][c] and base[r1][d] == base[r2][d]:
                                k[temp] += 1

                        A[r1][temp] = k[temp] / row
                        temp += 1
    print("done A")
    return A


def caculateM(base):
    colum = len(base[0])
    row = len(base)
    M = np.zeros((row, colum - 1))
    for t1 in range(row):
        k = [0] * (colum - 1)
        temp = 0
        for i in range(colum - 1):
            for t2 in range(row):
                if base[t1][i] == base[t2][i] and base[t1][colum - 1] == base[t2][colum - 1]:
                    k[temp] += 1
            M[t1][temp] = k[temp] / row
            temp += 1

    return M


def caculateB(base, A, M):
    colum = len(base[0])
    row = len(base)
    B = np.zeros((row, combination(3, colum - 1)))

    for r in range(row):
        temp = 0
        for a in range(0, colum - 3):
            for b in range(a + 1, colum - 2):
                for c in range(b + 1, colum - 1):
                    B[r][temp] = sum(A[r]) * min(M[r][a], M[r][b], M[r][c])
                    temp += 1
    print("done B")
    return B


def caculateC(base, B):
    colum = len(base[0])
    row = len(base)
    cols = 2 * combination(3, colum - 1)
    C = np.zeros((row, cols))

    for r1 in range(row):
        temp = 0
        for i in range(2):
            for a in range(0, (colum - 3)):
                for b in range(a + 1, (colum - 2)):
                    for c in range(b + 1, (colum - 1)):
                        for r2 in range(row):
                            if base[r1][a] == base[r2][a] and base[r1][b] == base[r2][b] and base[r1][c] == base[r2][c] and base[r2][colum - 1] == i:
                                C[r1][temp] += B[r2][temp % combination(3, colum - 1)]
                        temp += 1
    print("done C")
    return C


def writeToExcel(sheet, table, row):
    for x in range(row):
        for y in range(len(table[x])):
            sheet.cell(row=x + 1, column=y + 1, value=table[x][y])


def update(data):
    base = importFuzzyData(data)
    startA = time.time()
    A = caculateA(base)
    print("Time A: ", time.time() - startA)
    startB = time.time()
    M = caculateM(base)
    B = caculateB(base, A, M)
    print("Time B: ", time.time() - startB)
    startC = time.time()
    C = caculateC(base, B)
    print("Time C: ", time.time() - startC)
    writeToExcel(wb['A'], A, len(A))
    writeToExcel(wb['M'], M, len(M))
    writeToExcel(wb['B'], B, len(B))
    writeToExcel(wb['C'], C, len(C))

    wb.save("Data.xlsx")
    print("Update successful! Let's Start")


def FISA(base, C, list):
    colum = len(base[0])
    row = len(base)

    cols = combination(3, (colum - 1))
    C0 = [0] * cols
    C1 = [0] * cols

    t = 0
    for a in range(0, colum - 3):
        for b in range(a + 1, colum - 2):
            for c in range(b + 1, colum - 1):
                for r in range(row-1):
                    if base[r][a] == list[a] and base[r][b] == list[b] and base[r][c] == list[c] and base[r][colum-1] == 0:
                        C0[t] = C[r][t + 0 * cols]
                    if base[r][a] == list[a] and base[r][b] == list[b] and base[r][c] == list[c] and base[r][colum-1] == 1:
                        C1[t] = C[r][t + 1 * cols]
                t += 1

    D0 = max(C0) + min(C0)
    D1 = max(C1) + min(C1)

    if D0 > D1:
        return 0
    else:
        return 1

def Acc(A, B):
    result = 0
    valid_values = ['0', '1']  # Các giá trị hợp lệ mà A và B có thể có

    for i in range(len(A)):
        if A[i] in valid_values and B[i] in valid_values:
            if int(A[i]) == int(B[i]):
                result += 1

    return round(result * 100 / len(A), 2)



# def Tprecision(Pre,Act):
#     TP = 0
#     FP = 0
#
#     for i in range(len(Pre)):
#         if int(Pre[i]) == 1 and int(Act[i]) == 1:
#             TP +=1
#         if int(Pre[i]) == 1 and int(Act[i]) == 0:
#             FP +=1
#
#     if TP:
#         return round(100*TP/(TP+FP),2)
#     else:
#         if FP:
#             return 0
#         else:
#             return None

def Tprecision(Pre, Act):
    count_correct = 0
    count_total = 0

    for i in range(len(Pre)):
        try:
            if int(Pre[i]) == 1 and int(Act[i]) == 1:
                count_correct += 1
            if int(Act[i]) == 1:
                count_total += 1
        except ValueError:
            print(f"Skipping invalid data at index {i}: Pre={Pre[i]}, Act={Act[i]}")
            continue

    if count_total == 0:
        return 0
    return count_correct / count_total


def Trecall(Pre, Act):
    count_correct = 0
    count_total = 0

    for i in range(len(Pre)):
        try:
            if int(Pre[i]) == 1 and int(Act[i]) == 1:
                count_correct += 1
            if int(Act[i]) == 1:
                count_total += 1
        except ValueError:
            print(f"Skipping invalid data at index {i}: Pre={Pre[i]}, Act={Act[i]}")
            continue

    if count_total == 0:
        return 0
    return count_correct / count_total


listAcc = []
listPre = []
listRe = []
timeTest = []
timeUpdate = []

def testAccuracy(data, FuzzyTest2):
    test_data = importData(wb['FuzzyTest2'])
    sheet_C = wb['C']
    base = importFuzzyData(data)
    C = list(sheet_C.values)
    X = np.zeros(len(test_data))
    X_test = np.array(test_data).T[-1]

    for i in range(len(test_data)):
        try:
            X[i] = FISA(base, C, test_data[i])
        except ValueError:
            print(f"Invalid value encountered at index {i} in test data.")
            X[i] = 0  # Hoặc giá trị mặc định phù hợp

    listAcc.append(Acc(X, X_test))
    listPre.append(Tprecision(X, X_test))
    listRe.append(Trecall(X, X_test))

# def testAccuracy(data_sheet_name, test_sheet_name):
#     test_data = importData(wb[test_sheet_name])
#     X_test, Y_test = split_data(test_data)
#
#     X, Y = split_data(data)
#     listPre = []
#     listRe = []
#     listAcc = []
#
#     try:
#         listPre.append(Tprecision(X, X_test))
#         listRe.append(Trecall(X, X_test))
#         # Add similar handling for other functions if needed
#     except KeyError as e:
#         print(f"Error accessing worksheet: {e}")

    # Add other logic if needed

# for i in range(1, 21):
#     data_sheet_name = f"FuzzyData{i}"
#     test_sheet_name = f"FuzzyTest{i}"
#     print(f"Data sheet: {data_sheet_name}, Test sheet: {test_sheet_name}")
#     start = time.time()
#     update(data_sheet_name)
#     timeUpdate.append(round(time.time() - start, 2))
#
#     start2 = time.time()
#     testAccuracy(data_sheet_name, test_sheet_name)
#     timeTest.append(round(time.time() - start2, 2))

data_sheet_name = "FuzzyData2"
test_sheet_name = "FuzzyTest2"
print(f"Data sheet: {data_sheet_name}, Test sheet: {test_sheet_name}")
start = time.time()
update(data_sheet_name)
timeUpdate.append(round(time.time() - start, 2))

start2 = time.time()
testAccuracy(data_sheet_name, test_sheet_name)
timeTest.append(round(time.time() - start2, 2))

# Bạn có thể thực hiện các thao tác khác nếu cần


print(timeUpdate)
print(timeTest)
print(listAcc)

RESULTS = []
RESULTS.append(timeUpdate)
RESULTS.append(timeTest)
RESULTS.append(listAcc)
RESULTS.append(listPre)
RESULTS.append(listRe)
writeToExcel(wb['results'], RESULTS, len(RESULTS))
wb.save("Data.xlsx")
