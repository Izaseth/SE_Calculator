import pandas as pd
import string
import math


def main():
    excel = pd.read_excel('route_planner.xlsx')
    sum_of_S2(excel)
    S_sqrt(excel)
    sum_of_PT(excel)
    sum_of_LT(excel)
    sp_task(excel)
    excel.to_excel('route_planner.xlsx', index=False)

def S_sqrt(excel):
    S_values = excel["Path sum of squares"].values.tolist()
    for i in range(len(S_values)):
        value = math.sqrt(S_values[i])
        excel.loc[i, "Path Stdev"] = value

def sp_task(excel):
    nodes = excel["Task"].values.tolist()
    nodes.pop(0)
    index = 1
    for node in nodes:
        if isinstance(node, float):
            break
        filtered_excel = excel[excel["Final Node"] == node]
        sp_values = filtered_excel["Path Stdev"].tolist()
        excel.loc[index, "Node SP"] = max(sp_values)
        index+= 1



def sum_of_S2(excel):
    paths = excel['Path (one per row)'].values.tolist()
    task_value = dict(zip(excel['Task'], excel["S^2"]))
    sum_calculator(excel, paths, task_value,'Path sum of squares')

def sum_of_LT(excel):
    paths = excel['Path (one per row)'].values.tolist()
    task_value = dict(zip(excel['Task'], excel["LT"]))
    sum_calculator(excel, paths, task_value, 'LT Path')

def sum_of_PT(excel):
    paths = excel['Path (one per row)'].values.tolist()
    task_value = dict(zip(excel['Task'], excel["PT"]))
    sum_calculator(excel, paths, task_value, 'PT Path')

def sum_calculator(excel, paths, task_value, result_column:string):
    index_result_column = 0
    for path in paths:
        if isinstance(path, float):
            break
        value = 0
        tasks = path.split('-')
        for task in tasks:
            value += task_value[task]
        excel.loc[index_result_column, result_column] = value
        index_result_column += 1




if __name__ == '__main__':
    main()