import numpy as np
from scipy.optimize import linprog
from math import sqrt
from scipy import stats


def lp_array_1st_part(data, bestvalues, worstvalues, intervals, monotonicity):
    '''This function uses the elements of the bestvalues and worstvaleus and
    depending on the criteria monotonicity and the intervals the user wants,
    to create  a list with the discrete values of every criteria and with the
    array of data the uti;ities array for every alternative'''

    g = []
    for i in range(len(intervals)):
        gtemp = []
        a = bestvalues[i]
        b = worstvalues[i]
        c = intervals[i]
        step = (a-b)/c
        count_steps = intervals[i] + 1
        for j in range(count_steps):
            g_value = worstvalues[i]+step*j
            gtemp.append(g_value)
        g.append(gtemp)

    array_length = sum([len(g_elements) for g_elements in g])
    crit_dimension = len(data[0, :])
    alts_dimension = len(data[:, 0])
    utilities = np.zeros((alts_dimension, array_length))
    for j in range(alts_dimension):
        position = 0
        for i in range(crit_dimension):
            if monotonicity[i] == 1:
                for k in range(len(g[i])):
                    if data[j, i] == g[i][k]:
                        utilities[j, position] = 1

                    elif (data[j, i] > g[i][k]) and (data[j, i] < g[i][k+1]):
                        utilities[j, position] = (g[i][k+1] - data[j, i])/(g[i][k+1] - g[i][k])
                        utilities[j, position + 1] = (data[j, i]-g[i][k])/(g[i][k+1] - g[i][k])

                    position += 1
            if monotonicity[i] == 0:
                for k in range(len(g[i])):
                    if data[j, i] == g[i][k]:
                        utilities[j, position] = 1

                    elif (data[j, i] < g[i][k]) and (data[j, i] > g[i][k+1]):
                        utilities[j, position] = (g[i][k+1] - data[j, i])/(g[i][k+1] - g[i][k])
                        utilities[j, position +1] = (data[j, i]-g[i][k])/(g[i][k+1] - g[i][k])

                    position += 1

    for j in range(alts_dimension):
        position = 0
        for i in range(crit_dimension):
            if g[i][0] == data[j, i]:
                utilities[j, position] = 0
            position += len(g[i])

    return utilities, g


'''
#---------------------------test---------------------------
df = np.array([[3, 10, 1], [4, 20, 2], [2, 20, 0], [6, 40, 0], [30, 30, 3]])
bestvalues = [2, 10, 3]
worstvalues = [30, 40, 0]
monotonicity = [0, 0, 1]
intervals = [2, 3, 3]
ranking = [1, 2, 2, 3,  4]
utilities, g = lp_array_1st_part(df, bestvalues, worstvalues, intervals, monotonicity)
print(utilities)
#----------------------------------------------------------
'''


def lp_array_2nd_part_UTA(utilities, g, ranking, delta, sigma):
    '''
      This function creates the arrays of the linear program, it uses the differences of the rows
    of the utilities arrays to create the array of the equality restrictions and the upper bound restrictions
    of the linear program as well as the cost function
     utilities:
     g: The list with the discrete values for every criteri
     ranking: The list with the ranking of the alternatives
     delta: The difference between 2 alternatives
     sigma:
    return:
    '''
    equality_num = 0
    inequality_num = 0
    lp_rows = len(ranking)-1
    for i in range(lp_rows):
        if ranking[i] == ranking[i+1]:
            equality_num += 1
        else:
            inequality_num += 1
    alts_dimension = len(ranking)
    array_length = sum([len(g_elements) for g_elements in g])

    # lp_inequality is the A_ub array for the linprog function
    lp_inequality = np.zeros((inequality_num, array_length + alts_dimension))

    # lp_equality is the A_eq array for the linprog  function
    lp_equality = np.zeros((equality_num, array_length + alts_dimension))

    a = np.zeros((alts_dimension, alts_dimension))
    np.fill_diagonal(a, 1)
    utilities_new = np.concatenate((utilities, a), axis=1)

    eq_rows = 0
    ineq_rows = 0
    for i in range(lp_rows):
        if ranking[i] == ranking[i+1]:
            lp_equality[eq_rows, :] = utilities_new[i, :] - utilities_new[i+1, :]
            eq_rows += 1

        elif ranking[i] < ranking[i+1]:
            lp_inequality[ineq_rows, :] = utilities_new[i, :] - utilities_new[i+1, :]
            ineq_rows += 1

    # Because we want to create the differences for the g list for every sublist we will need
    # the quantity of the whole list g elements minus the quantity of sublists
    g_differences_rows = array_length - len(g)
    g_differences = np.zeros((g_differences_rows, array_length + alts_dimension))
    position_g_rows = 0
    position_g_columns = 0
    for i in range(len(g)):
        for j in range(len(g[i])-1):
            g_differences[position_g_rows, position_g_columns] = -1
            g_differences[position_g_rows, position_g_columns + 1] = 1
            position_g_rows += 1
            position_g_columns += 1
        position_g_columns += 1
    lp_inequality = np.concatenate((lp_inequality, g_differences), axis=0)

    lp_inequality = -1*lp_inequality


    g_rows = len(g)
    lb_eq_restriction = np.zeros((g_rows, array_length + alts_dimension))
    position_eq_lb = 0
    for i in range(g_rows):
        lb_eq_restriction[i, position_eq_lb] = 1
        position_eq_lb += len(g[i])

    lp_equality = np.concatenate((lp_equality, lb_eq_restriction), axis=0)

    ub_eq_restriction = np.zeros((1, array_length + alts_dimension))
    position_eq_ub = -1
    for i in range(g_rows):
        position_eq_ub += len(g[i])
        ub_eq_restriction[0, position_eq_ub] = 1

    lp_equality = np.concatenate((lp_equality, ub_eq_restriction), axis=0)

    rows_ineq = len(lp_inequality[:, 0])
    b_ineq = np.zeros(rows_ineq)
    for i in range(inequality_num):
        b_ineq[i] = -delta
    for i in range(inequality_num, len(b_ineq[:])):
        b_ineq[i] = -sigma

    rows_eq = len(lp_equality[:, 0])
    b_eq = np.zeros(rows_eq)
    b_eq[-1] = 1

    cost_function = np.zeros((1, array_length + alts_dimension))
    for i in range( array_length, array_length + alts_dimension):
        cost_function[0, i] = 1

    return cost_function, lp_inequality, lp_equality, b_ineq, b_eq


def cost_functionsUTA(data, c, g):
    '''Creates the cost functions for all the linear problems'''
    rows = len(data[0, :])
    columns = len(c[0, :])
    all_c = np.zeros((rows, columns))
    max_utilities = np.zeros((rows, columns))
    min_utilities = np.zeros((rows, columns))
    position = -1
    for i in range(len(g)):
        position += len(g[i])
        all_c[i, position] = 1

    return all_c, max_utilities, min_utilities


def calculate_utilitiesUTA(all_c, max_utilities, min_utilities, A_ineq, A_eq, b_ineq, b_eq, method):
    rows = len(all_c[:, 0])
    for i in range(rows):
        max_c = -1 * all_c[i, :]
        res = linprog(max_c, A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)
        max_utilities[i, :] = res.x

        min_c = all_c[i, :]
        res = linprog(min_c, A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)
        min_utilities[i, :] = res.x

    m1 = np.mean(max_utilities, axis=0)
    m2 = np.mean(min_utilities, axis=0)
    m3 = (m1 + m2) / 2

    mean_utilities = np.zeros((1, len(m3)))
    mean_utilities[0, :] = m3

    return mean_utilities, max_utilities, min_utilities


def lp_array_2nd_part_UTASTAR(data, utilities, g, ranking):
    w_rows = len(ranking)
    w_columns = sum([len(g_elements)-1 for g_elements in g])
    w = np.zeros((w_rows, w_columns))
    error = 0.000001
    for i in range(w_rows):
        pos_u = 0
        pos_w = 0

        for j in range(len(g)):
            corrupt = False
            if g[j][0] == data[i, j]:
                corrupt = True

            for k in range(len(g[j])-1):
                statement = utilities[i, pos_u] + utilities[i, pos_u+1]
                if statement < 1 - error:
                    w[i, pos_w] = 1
                    pos_w += 1
                    pos_u += 1
                elif statement > 1-error:
                    w[i, pos_w] = utilities[i, pos_u+1]
                    pos_w += len(g[j])-k-1
                    pos_u += len(g[j])-k
                    break

                if corrupt is True:
                    w[i, pos_w-1] = 0

            if corrupt is True:
                pos_u += 1

    return w


def create_error_array_Utastar(ranking):
    size = len(ranking)
    errors = np.zeros((size, 2*size))
    pos = 0
    for i in range(size):
        errors[i, pos] = 1
        errors[i, pos+1] = -1
        pos += 2
    return errors


def lp_array_3rd_part_UTASTAR(w, errors, ranking, delta):
    w_count = len(w[0, :])

    w = np.concatenate((w, errors), axis=1)
    columns = len(w[0, :])
    eq_rows = 0
    ineq_rows = 0
    alternatives = len(ranking)
    variable_count = len(w[0, :])
    for i in range(alternatives-1):
        if ranking[i] < ranking[i+1]:
            ineq_rows += 1
        elif ranking[i] == ranking[i+1]:
            eq_rows += 1

    A_eq = np.zeros((eq_rows, columns))
    A_ineq = np.zeros((ineq_rows, columns))
    eq_pos = 0
    ineq_pos = 0
    for i in range(alternatives-1):
        if ranking[i] == ranking[i+1]:
            A_eq[eq_pos, :] = w[i, :] - w[i+1, :]
            eq_pos += 1
        elif ranking[i] < ranking[i+1]:
            A_ineq[ineq_pos, :] = w[i, :] - w[i+1, :]
            ineq_pos += 1

    error_restriction = np.zeros((1, variable_count))
    for i in range(w_count):
        error_restriction[0, i] = 1

    cost_fun = np.zeros((1, variable_count))
    for i in range(w_count, variable_count):
        cost_fun[0, i] = 1

    A_eq = np.concatenate((A_eq, error_restriction), axis=0)
    A_ineq = -A_ineq

    b_eq = np.zeros((1, eq_rows+1))
    b_eq[0, -1] = 1

    b_ineq = np.full((1, ineq_rows), -delta)

    return cost_fun, A_eq, A_ineq, b_eq, b_ineq


def ASIcalculation(results, g, using_w):
    '''
    :param results: The results of the meta-optimization stage, including errors and transposition points
    :param g: A list that contains all the discrete values for every criteria
    :param using_w: If the method models the utilities through w transformation this parameter takes the value 1 else
                   it takes the value 0
    :return: The Average Stability Indices (ASI) for all the criteria
    If the method does not use w transformation counting excludes the zero values of the utilities
    s --> sum(from discrete value 1 to a-1)[ a*S(u^2)-(S(u))^2]
    '''
    breakpoint = sum([len(criterion_g)-using_w for criterion_g in g])

    #     ###    denominator construction   ###

    sum_a_minus_one = sum([len(criterion_g)-1 for criterion_g in g])
    denominator_1st_part = sum_a_minus_one
    sum_a = sum([len(criterion_g) for criterion_g in g])
    n_solutions = len(g)
    n_value = sum([i for i in range(len(g))])
    print(sum_a-n_value-n_solutions)
    denominator_2nd_part = 2*sqrt(abs(sum_a-n_value-n_solutions))
    denominator = denominator_1st_part*denominator_2nd_part

    #    ###    Sum Construction   ###

    Results = np.copy(results[:, 0: breakpoint])
    squaredResults = np.copy(np.square(Results))
    sum_of_utilities = np.sum(Results, axis=0)
    sum_of_squared_utilities = np.sum(squaredResults, axis=0)

    if using_w == 1:
        point = 0
    if using_w == 0:
        point = 1
    start_index = point
    end_index = 0
    numerator = 0

    for g_values in g:
        end_index += len(g_values) - using_w
        criterion_utilities_squares = np.copy(sum_of_squared_utilities[start_index:end_index])
        criterion_utilities = np.copy(sum_of_utilities[start_index:end_index])
        s = sum_a_minus_one*criterion_utilities_squares-np.square(criterion_utilities)
        sqrt_s = np.sqrt(abs(s))
        sum_of_sqrt_s = np.sum(sqrt_s)
        numerator += sum_of_sqrt_s
        start_index = end_index + point

    ASI = 1-numerator/denominator

    return ASI


def calculate_utilitiesUTASTAR(g, A_ineq, A_eq, b_eq, b_ineq, method):
    variables = len(A_ineq[0, :])
    funcs_count = len(g)
    cost_funcs = np.zeros((funcs_count, variables))
    pos = 0
    for i in range(funcs_count):
        for j in range(len(g[i])-1):
            cost_funcs[i, pos] = -1
            pos += 1

    results = np.zeros((funcs_count, variables))
    for i in range(funcs_count):
        res = linprog(cost_funcs[i, :], A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)
        results[i, :] = res.x

    mean_util = np.mean(results, axis=0)
    return mean_util, results


def add_probabilities(w_array, data, monotonicity, distributions, g, metrics):
    row_search = len(data[:,0])
    column_search = len(data[0, :])
    prob = np.zeros((row_search,column_search))
    for i in range(row_search):
        for j in range(column_search):
            if monotonicity[j] is 0:
                uni_loc = g[j][-1]
                uni_scale = g[j][0]-g[j][-1]
                norm_loc = metrics[j, 0]
                p = calculate_probability(data[i, j], distributions[j], monotonicity[j], norm_loc=norm_loc, norm_scale=metrics[j, 1],uni_loc=uni_loc, uni_scale=uni_scale)
                prob[i, j] = 1 - p
            if monotonicity[j] is 1:
                uni_loc = g[j][0]
                uni_scale = g[j][-1] - g[j][0]
                norm_loc = metrics[j, 0]
                p = calculate_probability(data[i, j], distributions[j], monotonicity[j],norm_loc=norm_loc, norm_scale=metrics[j, 1], uni_loc=uni_loc, uni_scale=uni_scale)
                prob[i, j] = p
    breaking_points = [len(g_criterion)-1 for g_criterion in g]

    wp_array = np.zeros((len(w_array[:, 0]), len(w_array[0, :])))
    pos = 0
    for i in range(len(breaking_points)):
        for j in range(breaking_points[i]):
            wp_array[:, pos] = w_array[:, pos]*prob[:, i]
            pos += 1

    return wp_array, prob

def utility_criterion(data, crit_utilities, g, monotonicity):

    crit_dimension = len(data[0, :])
    alts_dimension = len(data[:, 0])
    p_utilities = np.zeros((alts_dimension, crit_dimension))
    for j in range(alts_dimension):
        position = 0
        for i in range(crit_dimension):
            if monotonicity[i] == 1:
                for k in range(len(g[i])):
                    if data[j, i] == g[i][k]:
                        p_utilities[j, i] = crit_utilities[position]

                    elif (data[j, i] > g[i][k]) and (data[j, i] < g[i][k + 1]):
                        a = crit_utilities[position]*((g[i][k + 1] - data[j, i]) / (g[i][k + 1] - g[i][k]))
                        b = crit_utilities[position + 1]*( (data[j, i] - g[i][k]) / (g[i][k + 1] - g[i][k]))
                        p_utilities[j, i] = a+b

                    position += 1
            if monotonicity[i] == 0:
                for k in range(len(g[i])):
                    if data[j, i] == g[i][k]:
                        p_utilities[j, i] = crit_utilities[position]

                    elif (data[j, i] < g[i][k]) and (data[j, i] > g[i][k + 1]):
                        a = crit_utilities[position]*((g[i][k + 1] - data[j, i]) / (g[i][k + 1] - g[i][k]))
                        b = crit_utilities[position + 1]*((data[j, i] - g[i][k]) / (g[i][k + 1] - g[i][k]))
                        p_utilities[j, i] = a+b

                    position += 1
    return p_utilities


def calculate_probability(x, distribution,monotonicity, norm_loc=None, norm_scale=None, uni_loc=None, uni_scale=None):
    if distribution is 'uniform':
        prob = stats.uniform.cdf(x, loc=uni_loc, scale=uni_scale)
        return prob
    if distribution is 'normal':
        prob = stats.norm.cdf(x, loc=norm_loc, scale=norm_scale)
        return prob
    if distribution is 'det':
        if monotonicity is 0:
            return 0
        if monotonicity is 1:
            return 1


def utility_aggregation(mean_util, g):
    utils_count = len(mean_util[:]) + len(g)
    aggregated_utils = np.zeros((1, utils_count))
    cu_pos = 0
    mu_pos = 0
    for i in range(len(g)):
        sum = 0
        cu_pos += 1
        for j in range(len(g[i])-1):
            sum += mean_util[mu_pos]
            aggregated_utils[0, cu_pos] = sum
            mu_pos += 1
            cu_pos += 1
    aggregated_utils[0,cu_pos:] = mean_util[mu_pos:]
    return aggregated_utils


def answers1(initial_utils, g, utilities, alternatives):
    crit_u_names = []
    crit_index = 1
    subcrit_index = 1
    for i in range(len(g)):
        for j in range(len(g[i])):
            u = "u" + str(crit_index) + "." + str(subcrit_index)
            crit_u_names.append(u)
            subcrit_index += 1
        crit_index += 1
        subcrit_index = 1

    break_point = len(crit_u_names)
    crit_utilities = initial_utils[0, 0: break_point]
    errors = initial_utils[0, break_point:]

    length = len(alternatives)
    alts_utilities = np.zeros(length)
    for i in range(length):
        alts_utilities[i] = np.sum(utilities[i, :] * crit_utilities)

    return alts_utilities, crit_utilities, crit_u_names, errors


def answers2(mean_utilities, intervals):
    breakpoint = sum([i for i in intervals])
    w = np.zeros((1,breakpoint))
    w[0,:] = mean_utilities[:breakpoint]
    w_names = []
    crit_index = 1
    for i in intervals:
        for subcrit_index in range(1,i+1):
            w_name = 'w' + str(crit_index) + '.' + str(subcrit_index)
            w_names.append(w_name)
        crit_index += 1

    return w, w_names


def ktaumetric(alts_utilities, ranking):
    autilities = 1 / (alts_utilities + 0.0000000001)
    uranking = [sorted(autilities).index(x)+1 for x in autilities]
    tau, p_value = stats.kendalltau(autilities, ranking)
    return tau
'''
############################## TEST ##############################
df = np.array([[3, 10, 1], [4, 20, 2], [2, 20, 0], [6, 40, 0], [30, 30, 3]])
bestvalues = [2, 10, 3]
worstvalues = [30, 40, 0]
monotonicity = [0, 0, 1]
intervals = [2, 3, 3]
ranking = [1, 2, 2, 3,  4]
utilities, g = lp_array_1st_part(df, bestvalues, worstvalues, intervals, monotonicity)
print(utilities)
lp_array_2nd_part_UTASTAR(df, utilities, g, ranking, delta=0.01, sigma=0)
##################################################################
'''

