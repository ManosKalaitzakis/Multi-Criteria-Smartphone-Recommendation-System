import numpy as np
from scipy.optimize import linprog


def plus_error(A_ineq,w_values, pos_error, count_w, i, pos_a, pos_u):
    A_ineq[pos_a, 0:count_w] = -w_values[i, :]
    A_ineq[pos_a, pos_error] = -1
    A_ineq[pos_a, pos_u] = 1
    return A_ineq


def minus_error(A_ineq,w_values, pos_error, count_w, i, pos_a, pos_u, b_ineq, delta):
    A_ineq[pos_a, 0:count_w] = w_values[i, :]
    A_ineq[pos_a, pos_error] = -1
    A_ineq[pos_a, pos_u] = -1
    b_ineq[pos_a,0] = -delta
    return A_ineq, b_ineq


def get_g(bestvalues, worstvalues, intervals):
    g = []
    size = len(bestvalues)
    for i in range(size):
        step = (bestvalues[i]-worstvalues[i])/intervals[i]
        temp_list = [worstvalues[i]]
        temp = worstvalues[i]
        for j in range(intervals[i]):
            temp = temp + step
            temp_list.append(temp)
        g.append(temp_list)
    return g

def get_u_part(data, g, monotonicity):
    u_length = sum([len(g_elements)-1 for g_elements in g])
    u_array = np.zeros((len(data[:,0]), u_length ))
    for i in range(len(data[:, 0])):
        pos_u = 0
        for j in range(len(g)):
            if monotonicity[j] == 1:
                for k in range(1, len(g[j])):
                    if data[i,j] >= g[j][k]:
                        u_array[i, pos_u] = 1
                    elif g[j][k-1] < data[i, j] < g[j][k]:
                        u_array[i, pos_u] = (data[i, j]-g[j][k-1])/(g[j][k]-g[j][k-1])
                    pos_u += 1
            elif monotonicity[j] == 0:
                for k in range(1, len(g[j])):
                    if data[i,j] <= g[j][k]:
                        u_array[i, pos_u] = 1
                    elif g[j][k-1] > data[i, j] > g[j][k]:
                        u_array[i, pos_u] = (data[i, j]-g[j][k-1])/(g[j][k]-g[j][k-1])
                    pos_u += 1
    return u_array


def linprog_dis_inequalities(u_values, groups, pref_order, delta, sigma):
#### Part 1 define the size of the A_ub array  ####
    count_errors = 0
    for i in range(len(groups)):
        if groups[i] == pref_order[0] or groups[i] == pref_order[-1]:
            count_errors += 1
        else: 
            count_errors += 2
    count_u = len(pref_order)-1
    count_w = len(u_values[0, :])
    col_num = count_w + count_u + count_errors
    row_num = count_errors
    A_ineq = np.zeros((row_num, col_num))
    b_ineq = np.zeros((row_num, 1))

#### Part 2 fill the A_ub array  ####
    pos_error = count_w + count_u
    pos_a = 0
    for i in range(len(groups)):
        if groups[i] == pref_order[0]:
            pos_u = count_w+count_u-1
            A_ineq, b_ineq = minus_error(A_ineq, u_values, pos_error, count_w, i, pos_a, pos_u, b_ineq, delta)
            pos_error += 1
            pos_a += 1
        if groups[i] == pref_order[-1]:
            pos_u = count_w
            A_ineq = plus_error(A_ineq, u_values, pos_error, count_w, i, pos_a, pos_u)
            pos_error += 1
            pos_a += 1
        counter = len(pref_order)-1
        for j in range(1, counter):
            if groups[i] == pref_order[j]:
                pos_u = count_w + j-1
                A_ineq, b_ineq = minus_error(A_ineq, u_values, pos_error, count_w, i, pos_a, pos_u, b_ineq, delta)
                pos_error += 1
                pos_a += 1
                
                pos_u = count_w + j
                A_ineq = plus_error(A_ineq, u_values, pos_error, count_w, i, pos_a, pos_u)
                pos_error += 1
                pos_a += 1

#### Part 3 adding the last restrictions of group preferency ####
    if len(pref_order) >= 3:
        rows_u = len(pref_order)-2
        cols_u = len(A_ineq[0, :])
        u_rest = np.zeros((rows_u, cols_u))
        bIneq_rest = np.full((rows_u, 1), -sigma)
        pos_u = count_w
        for i in range(len(u_rest[:, 0])):
            u_rest[i, pos_u] = -1
            u_rest[i, pos_u + 1] = 1
            pos_u += 1
        A_ineq = np.concatenate((A_ineq, u_rest))
        b_ineq = np.concatenate((b_ineq, bIneq_rest))
    return A_ineq, b_ineq


def linprog_dis_equalities(A_ineq, g):
    A_rows = 1
    A_eq = np.zeros((A_rows, len(A_ineq[0, :])))
    pos_a1 = 0
    for i in range(len(g)):
        for j in range(1,len(g[i])):
            A_eq[0, pos_a1] = 1
            pos_a1 += 1

    b_eq = np.ones((1, 1))

    return A_eq, b_eq


def linprog_dis_costfunction(A_ineq, u_values, pref_order):
    c = np.ones((1, len(A_ineq[0, :])))
    stop = len(u_values[0, :]) + len(pref_order) - 1
    for i in range(stop): c[0, i] = 0

    return c


def meta_optimization(res, c, A_ineq, b_ineq, A_eq, b_eq, g, method, epsilon):
    '''
    min_c: the cost functions for minimization
    max_c: the cost functions for maximization
    '''
    A_ineq = np.concatenate((A_ineq, c))
    b_ineq = np.append(b_ineq, res.fun+epsilon)
    min_c = np.zeros((len(g), len(A_ineq[0, :])))
    pos = 0
    for i in range(len(g)):
        for j in range(1, len(g[i])):
            min_c[i, pos] = 1
            pos += 1
    max_c = - min_c
    max_x = np.zeros((len(max_c), len(max_c[0, :])))
    min_x = np.zeros((len(min_c), len(min_c[0, :])))
    for i in range(len(max_c)):
        resmeta = linprog(max_c[i, :], A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)
        max_x[i, :] = resmeta.x

        resmeta = linprog(min_c[i, :], A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)
        min_x[i, :] = resmeta.x

    m1 = np.mean(max_x, axis=0)
    m2 = np.mean(min_x, axis=0)
    m3 = (m1 + m2) / 2

    mean_utilities = m3

    return mean_utilities, max_x, min_x


def x_deconstruction(mean_utilities, g, pref_order):

    '''
    This function breaks the mean_utilities array in 3 parts:
    w_values: This array will be used to calculate the predicting values.
    transitionpoints: The values of this array will be used to classify the alternatives of the
    unknown data
    errors: This array can be used by the analyst to have an insight on the results
    '''

    breakpoint1 = sum([len(l)-1 for l in g])
    w_values = mean_utilities[:breakpoint1]

    breakpoint2 = breakpoint1 + len(pref_order) - 1
    transitionpoints = mean_utilities[breakpoint1:breakpoint2]

    errors = mean_utilities[breakpoint2:]

    return w_values, transitionpoints, errors


def estimate_values(res, useful, data):
    alts_count = len(data)
    estimation = np.zeros(alts_count)
    for row in range(alts_count):
        pos_w = 0
        alt_estimation = 0
        for i in range(len(res.g)):
            if useful.monotonicity[i] == 1:
                for j in range(len(res.g[i])-1):
                    if data[row, i] >= res.g[i][j+1]:
                        alt_estimation += res.w_values[pos_w]
                    elif res.g[i][j] < data[row, i] < res.g[i][j+1]:
                        alt_estimation += res.w_values[pos_w]*((data[row, i]-res.g[i][j])/(res.g[i][j+1]-res.g[i][j]))
                    pos_w += 1
            elif useful.monotonicity[i] == 0:
                for j in range(len(res.g[i])-1):
                    if data[row, i] <= res.g[i][j+1]:
                        alt_estimation += res.w_values[pos_w]
                    elif res.g[i][j] > data[row, i] > res.g[i][j+1]:
                        alt_estimation += res.w_values[pos_w]*((data[row, i]-res.g[i][j])/(res.g[i][j+1]-res.g[i][j]))
                    pos_w += 1
        estimation[row] = alt_estimation

    return estimation


def predict_classes(res, useful, estimation):
    '''This method predicts in which class the alternative belongs '''
    prediction_list = []
    for est in estimation:
        index = len(useful.pref_order)-1
        for i in range(len(res.transitionpoints)):
            if est > res.transitionpoints[i]:
                prediction_list.append(useful.pref_order[index])
                break
            index -= 1
        if est <= res.transitionpoints[i]:
            prediction_list.append(useful.pref_order[0])

    return prediction_list


def w_aggregation(w_values, g):
    length = sum([len(g_elements) for g_elements in g])
    crit_utilities = np.zeros(length)
    indexcu = 0
    indexw = 0
    for gList in g:
        indexcu += 1
        w_aggregation = 0
        for i in range(len(gList)-1):
            w_aggregation += w_values[indexw]
            crit_utilities[indexcu] = w_aggregation
            indexw += 1
            indexcu += 1

    return crit_utilities





