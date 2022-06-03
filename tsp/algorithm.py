import pandas as pd
import copy
import sys

if __name__ == '__main__':
    #loading of data
    val = input("Enter your excel sheet path please: ")
    if val == '':
        print('please enter a valid path thank you')
        quit()
    a = ".xlsx"
    valf = val + a
    try:
        df = pd.read_excel(valf, header=None)
    except KeyError as e:
        print("Expected column headers not found")
        sys.exit(1)
    except TypeError as e:
        print("Type Error")
        sys.exit(1)
    except FileNotFoundError as e:
        print("Excel file not found " + str(e))
        sys.exit(1)
    destination = (df.iloc[-1]).to_dict()
    destination.popitem()
    sourc = (df.iloc[:, -1]).to_dict()
    sourc.popitem()
    source = []
    for i, j in sourc.items():
        source.append(j)

    total = []

    df_val = (df.iloc[:, 0:-1]).to_dict('list')
    values = []
    for i, j in df_val.items():
        values.append(j)
    original = copy.deepcopy(values)

    des_values = destination.values()
    des_sum = sum(des_values)
    sou_values = sourc.values()
    sou_sum = sum(sou_values)
    #test of condition of balance
    if des_sum == sou_sum:
        #testing of available quantity
        while (sum(source) > 0):
            # algorithm core
            val_min = min(x if isinstance(x, int) else min(x) for x in values)
            # get nested lists iwth min val
            multi = [x for x in values if val_min in x]
            multi_min = 9999999999
            best = []
            for i in multi:
                if multi_min > i[-1]:
                    multi_min = i[-1]
                    best.clear()
                    best.append(i)
                    best.append(values.index(i))
                    best.append(i.index(val_min))

            for item in values:
                for i in item[:-1]:
                    iindex = item.index(i)
                    if i != val_min:
                        item[iindex] -= val_min

            if values[best[1]][-1] <= source[best[2]]:
                total.append((original[best[1]][best[2]], values[best[1]][-1]))
                source[best[2]] -= values[best[1]][-1]
                values.pop(best[1])
                original.pop(best[1])
            else:
                total.append((original[best[1]][best[2]], source[best[2]]))
                for i in values:
                    del i[best[2]]

                for j in original:
                    del j[best[2]]

                values[best[1]][-1] -= source[best[2]]
                source.pop(best[2])

        cost = 0
        total = tuple((int(x[0]), x[1]) for x in total)
        for i in total:
            cost += i[0] * i[1]
        print('The total Transportation cost is:', total, '=', cost)
    else:
        print('please balace your systeme')
