import pandas as pd
from filters import CheckData, ObjectDictionary, OptimizeResult, sort
from utautastarfunctions import *
from utadisfunctions import *
from scipy.optimize import linprog
from matplotlib import pyplot as plt
from scipy.stats import kendalltau
from scipy import stats
import pickle


def UTA(data, alternatives=None, criteria=None, ranking=None, monotonicity=None,
        bestvalues=None, worstvalues=None, intervals=None, method='simplex',
        epsilon=0.009, delta=0.01, sigma=0.01):

    filter = CheckData(data)
    if isinstance(data, pd.DataFrame):
        filter.ranking_pd()
        filter.reposition_ranking_pd()
        filter.filter_values_pd()
        ranking = filter.cut_the_ranking()
        data, criteria, alternatives = filter.turn_to_np()
    elif isinstance(data, np.ndarray):
        alternatives = filter.list_filter(alternatives, 0, "alternative")
        criteria = filter.list_filter(criteria, 1, "criterion")
        ranking = filter.ranking_np(ranking)
        ranking, data = filter.filter_values_np(ranking)

    else:
        raise ValueError("This form of data is not supported by the function")

    monotonicity = filter.gradient(monotonicity)
    bestvalues = filter.BestValues(bestvalues, monotonicity)
    worstvalues = filter.WorstValues(worstvalues, monotonicity)
    intervals = filter.Intervals(intervals)

    utilities_multiplier, g = lp_array_1st_part(data, bestvalues, worstvalues, intervals, monotonicity)

    c, A_ineq, A_eq, b_ineq, b_eq = lp_array_2nd_part_UTA(utilities_multiplier, g, ranking, delta, sigma)

    res = linprog(c, A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)

    A_eq = np.concatenate((A_eq, c), axis=0)
    b_eq = np.append(b_eq, res.fun + epsilon)

    all_c, max_utilities, min_utilities = cost_functionsUTA(data, c, g)

    mean_utilities, max_utilities, min_utilities = calculate_utilitiesUTA(all_c, max_utilities, min_utilities, A_ineq, A_eq, b_ineq, b_eq, method)

    results = np.concatenate((max_utilities, min_utilities), axis=0)

    using_w = 0

    breakpoint = sum([len(criterion_g) - using_w for criterion_g in g])
    Results = np.copy(results[:, 0: breakpoint])

    ASI = ASIcalculation(results, g, using_w)

    alts_utilities, crit_utilities, crit_names, errors = answers1(mean_utilities, g, utilities_multiplier, alternatives)

    tau = ktaumetric(alts_utilities, ranking)

    results = {
        'g': g,
        'criteria': criteria,
        'crit_names': crit_names,
        'crit_utilities': crit_utilities,
        'all_utilities': Results,
        'alternatives': alternatives,
        'alts_utilities': alts_utilities,
        'errors': errors,
        'tau': tau,
        'ASI': ASI}

    return OptimizeResult(results)


def UTASTAR(data, alternatives=None, criteria=None, ranking=None, monotonicity=None, bestvalues=None,
            worstvalues=None, intervals=None,  method='simplex', epsilon=0.001, delta=0.05):
    filter = CheckData(data)
    if isinstance(data, pd.DataFrame):
        filter.ranking_pd()
        filter.reposition_ranking_pd()
        filter.filter_values_pd()
        ranking = filter.cut_the_ranking()
        data, criteria, alternatives = filter.turn_to_np()
    elif isinstance(data, np.ndarray):
        alternatives = filter.list_filter(alternatives, 0, "alternative")
        criteria = filter.list_filter(criteria, 1, "criterion")
        ranking = filter.ranking_np(ranking)
        ranking, data = filter.filter_values_np(ranking)

    else:
        raise ValueError("This form of data is not supported by the function")

    monotonicity = filter.gradient(monotonicity)
    bestvalues = filter.BestValues(bestvalues, monotonicity)
    worstvalues = filter.WorstValues(worstvalues, monotonicity)
    intervals = filter.Intervals(intervals)

    utilities_multiplier, g = lp_array_1st_part(data, bestvalues, worstvalues, intervals, monotonicity)
    w_array = lp_array_2nd_part_UTASTAR(data, utilities_multiplier, g, ranking)
    errors_array = create_error_array_Utastar(ranking)
    c, A_eq, A_ineq, b_eq, b_ineq = lp_array_3rd_part_UTASTAR(w_array, errors_array, ranking, delta)

    res = linprog(c, A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)

    A_eq = np.concatenate((A_eq, c), axis=0)
    b_eq = np.append(b_eq, res.fun + epsilon)

    mean_util, results = calculate_utilitiesUTASTAR(g, A_ineq, A_eq, b_eq, b_ineq, method)

    using_w = 1

    breakpoint = sum([len(criterion_g) - using_w for criterion_g in g])
    Results = np.copy(results[:, 0: breakpoint])

    ASI = ASIcalculation(results, g, using_w)

    aggregated_utils = utility_aggregation(mean_util, g)

    alts_utilities, crit_utilities, crit_names, errors = answers1(aggregated_utils, g, utilities_multiplier, alternatives)

    w, w_names = answers2(mean_util, intervals)

    tau = ktaumetric(alts_utilities, ranking)

    results = {
        'g': g,
        'criteria': criteria,
        'crit_names': crit_names,
        'crit_utilities': crit_utilities,
        'w_names': w_names,
        'w': w,
        'all_w': Results,
        'alternatives': alternatives,
        'alts_utilities': alts_utilities,
        'errors': errors,
        'tau': tau,
        'ASI': ASI}

    return OptimizeResult(results)


def UTAII(data, alternatives=None, criteria=None, ranking=None, monotonicity=None,
          bestvalues=None, worstvalues=None, intervals=None, method='simplex',
          epsilon=0.001, delta=0.05, sigma=0.01, utilities='UTA'):
    if utilities is 'UTA':
        res = UTA(data, alternatives=alternatives, criteria=criteria, ranking=ranking, monotonicity=monotonicity,
                  bestvalues=bestvalues, worstvalues=worstvalues, intervals=intervals, method=method,
                  epsilon=epsilon, delta=delta, sigma=sigma)
    elif utilities is 'UTASTAR':
        res = UTASTAR(data, alternatives=alternatives, criteria=criteria, ranking=ranking, monotonicity=monotonicity,
                      bestvalues=bestvalues, worstvalues=worstvalues, intervals=intervals, method=method,
                      epsilon=epsilon, delta=delta)
    else:
        raise ValueError('Not acceptable utility calculation method')

    filter = CheckData(data)
    if isinstance(data, pd.DataFrame):
        filter.ranking_pd()
        filter.reposition_ranking_pd()
        filter.filter_values_pd()
        ranking = filter.cut_the_ranking()
        data, criteria, alternatives = filter.turn_to_np()
    elif isinstance(data, np.ndarray):
        alternatives = filter.list_filter(alternatives, 0, "alternative")
        criteria = filter.list_filter(criteria, 1, "criterion")
        ranking = filter.ranking_np(ranking)
        ranking, data = filter.filter_values_np(ranking)

    else:
        raise ValueError("This form of data is not supported by the function")

    monotonicity = filter.gradient(monotonicity)
    bestvalues = filter.BestValues(bestvalues, monotonicity)
    worstvalues = filter.WorstValues(worstvalues, monotonicity)
    intervals = filter.Intervals(intervals)

    g = res.g
    crit_utilities = res.crit_utilities

    p_utilities = utility_criterion(data, crit_utilities, g, monotonicity)

    array_size = len(p_utilities[:, 0])
    errors_array = np.zeros([array_size, array_size*2])
    col_index = 0
    for row in range(array_size):
        errors_array[row, col_index] = 1
        errors_array[row, col_index+1] = -1
        col_index += 2

    cost_fun, A_eq, A_ineq, b_eq, b_ineq = lp_array_3rd_part_UTASTAR(p_utilities, errors_array, ranking, delta)

    res1 = linprog(cost_fun, A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)

    A_eq = np.concatenate((A_eq, cost_fun), axis=0)
    b_eq = np.append(b_eq, res1.fun + epsilon)

    variables = len(A_ineq[0, :])
    funcs_count = len(p_utilities[0, :])
    cost_funcs = np.zeros((funcs_count, variables))
    for i in range(funcs_count):
        cost_funcs[i, i] = -1

    results = np.zeros((funcs_count, variables))
    for i in range(funcs_count):
        res = linprog(cost_funcs[i, :], A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)
        results[i, :] = res.x

    mean_util = np.mean(results, axis=0)

    using_w =0

    breakpoint = sum([len(criterion_g) - using_w for criterion_g in g])
    Results = np.copy(results[:, 0: breakpoint])

    ASI = ASIcalculation(results, g, using_w)

    crit_names = []
    crit_index = 1
    subcrit_index = 1
    for i in range(len(g)):
        for j in range(len(g[i])):
            u = "u" + str(crit_index) + "." + str(subcrit_index)
            crit_names.append(u)
            subcrit_index += 1
        crit_index += 1
        subcrit_index = 1

    break_point = len(data[0, :])
    p_values = mean_util[0: break_point]
    errors = mean_util[break_point:]

    length = len(alternatives)
    alts_utilities = np.zeros(length)
    for i in range(length):
        alts_utilities[i] = np.sum(p_utilities[i, :] * p_values)

    tau = ktaumetric(alts_utilities, ranking)

    results = {
        'g': g,
        'criteria': criteria,
        'crit_names': crit_names,
        'crit_utilities': crit_utilities,
        'p': p_values,
        'all_p': Results,
        'alternatives': alternatives,
        'alts_utilities': alts_utilities,
        'errors': errors,
        'tau': tau,
        'ASI': ASI}

    return OptimizeResult(results)


def stochastic_UTA(data, alternatives=None, criteria=None, ranking=None, monotonicity=None, bestvalues=None,
                   worstvalues=None, distributions=None, metrics=None, intervals=None, epsilon=0.005, delta=0.05, method='simplex'):
    filter = CheckData(data)
    if isinstance(data, pd.DataFrame):
        filter.ranking_pd()
        filter.reposition_ranking_pd()
        filter.filter_values_pd()
        ranking = filter.cut_the_ranking()
        data, criteria, alternatives = filter.turn_to_np()
    elif isinstance(data, np.ndarray):
        alternatives = filter.list_filter(alternatives, 0, "alternative")
        criteria = filter.list_filter(criteria, 1, "criterion")
        ranking = filter.ranking_np(ranking)
        ranking, data = filter.filter_values_np(ranking)

    else:
        raise ValueError("This form of data is not supported by the function")

    monotonicity = filter.gradient(monotonicity)
    bestvalues = filter.BestValues(bestvalues, monotonicity)
    worstvalues = filter.WorstValues(worstvalues, monotonicity)
    intervals = filter.Intervals(intervals)

    acceptable_distributions = ['normal', 'uniform', 'det']

    for distribution in distributions:
        if distribution not in acceptable_distributions:
            raise ValueError('Not acceptable distribution')

    utilities_multiplier, g = lp_array_1st_part(data, bestvalues, worstvalues, intervals, monotonicity)
    w_array = lp_array_2nd_part_UTASTAR(data, utilities_multiplier, g, ranking)

    wp_array, prob = add_probabilities(w_array, data, monotonicity, distributions, g, metrics)

    errors_array = create_error_array_Utastar(ranking)
    c, A_eq, A_ineq, b_eq, b_ineq = lp_array_3rd_part_UTASTAR(wp_array, errors_array, ranking, delta)

    res = linprog(c, A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)

    A_eq = np.concatenate((A_eq, c), axis=0)
    b_eq = np.append(b_eq, res.fun + epsilon)

    mean_util, results = calculate_utilitiesUTASTAR(g, A_ineq, A_eq, b_eq, b_ineq, method)

    using_w = 1

    breakpoint = sum([len(criterion_g) - using_w for criterion_g in g])
    Results = np.copy(results[:, 0: breakpoint])

    ASI = ASIcalculation(results, g, using_w)

    aggregated_utils = utility_aggregation(mean_util, g)

    useless, crit_utilities, crit_names, errors = answers1(aggregated_utils, g, utilities_multiplier,
                                                                  alternatives)
    p_utilities = utility_criterion(data, crit_utilities, g, monotonicity)

    w, w_names = answers2(mean_util, intervals)

    length = len(alternatives)
    alts_utilities = np.zeros(length)
    for i in range(length):
        alts_utilities[i] = np.sum(p_utilities[i, :] * prob[i,:])

    tau = ktaumetric(alts_utilities, ranking)

    results = {
        'g': g,
        'criteria': criteria,
        'crit_names': crit_names,
        'crit_utilities': crit_utilities,
        'w_names': w_names,
        'w': w,
        'prob_array': prob,
        'all_w': Results,
        'alternatives': alternatives,
        'alts_utilities': alts_utilities,
        'errors': errors,
        'tau': tau,
        'ASI': ASI}

    return OptimizeResult(results)


class UTADIS():
    ''' UTADIS is a multicriteria method used for creating an additive utility function in order to classify
    alternatives into predefined strict groups of preference
    '''

    def __init__(self):
        self.optresults = None
        self.useful = None
        self.predictions = None
        self.estimation = None

    def fit(self, data, groups, pref_order, alternatives=None, criteria=None, monotonicity=None, bestvalues=None,
            worstvalues=None, intervals=None, epsilon=0.005, delta=0.05, sigma=0.1,
            method='simplex'):
        ''' Use data to create the function that contains the additive utility function, the border
        of the groups and the errors '''
        filter = CheckData(data)
        if isinstance(data, pd.DataFrame):
            filter.filter_values_pd()
            data, criteria, alternatives = filter.turn_to_np()
        elif isinstance(data, np.ndarray):
            alternatives = filter.list_filter(alternatives, 0, "alternative")
            criteria = filter.list_filter(criteria, 1, "criterion")
            data = filter.filter_nan_UTADIS()
        else:
            raise ValueError("This form of data is not supported")

        monotonicity = filter.gradient(monotonicity)
        bestvalues = filter.BestValues(bestvalues, monotonicity)
        worstvalues = filter.WorstValues(worstvalues, monotonicity)
        intervals = filter.Intervals(intervals)
        filter.group_consistency(groups, pref_order)

        g = get_g(bestvalues, worstvalues, intervals)

        u_values = get_u_part(data, g, monotonicity)

        A_ineq, b_ineq = linprog_dis_inequalities(u_values, groups, pref_order, delta, sigma)

        A_eq, b_eq = linprog_dis_equalities(A_ineq, g)

        c = linprog_dis_costfunction(A_ineq, u_values, pref_order)

        res = linprog(c, A_ub=A_ineq, b_ub=b_ineq, A_eq=A_eq, b_eq=b_eq, method=method)

        mean_utilities, max_x, min_x = meta_optimization(res, c, A_ineq, b_ineq, A_eq, b_eq, g, method, epsilon)

        results = np.concatenate((max_x, min_x), axis=0)

        using_w = 1

        breakpoint = sum([len(criterion_g) - using_w for criterion_g in g])
        Results = np.copy(results[:, 0: breakpoint])

        ASI = ASIcalculation(results, g, using_w)

        w_values, transitionpoints, errors = x_deconstruction(mean_utilities, g, pref_order)

        crit_utilities = w_aggregation(w_values, g)

        useful = {
            'pref_order': pref_order,
            'monotonicity': monotonicity}

        self.useful = ObjectDictionary(useful)

        results = {
            'g': g,
            'w_values': w_values,
            'all_w': Results,
            'transitionpoints': transitionpoints,
            'errors': errors,
            'crit_utilities': crit_utilities,
            'criteria': criteria,
            'ASI': ASI}
        self.optresults = OptimizeResult(results)

        return self.optresults

    def predict(self, data):
        '''This method predicts the classes for a new set of alternatives with the w values
        of the latest data set.
        data: Numpy array that contains the alternatives with the unknown class.
              The values of every criteria mus be in between the range of best values
              and worst values'''
        filter = CheckData(data)
        if isinstance(data, pd.DataFrame):
            filter.filter_values_pd()
            data, criteria, alternatives = filter.turn_to_np()
        elif isinstance(data, np.ndarray):
            data = filter.filter_nan_UTADIS()
        else:
            raise ValueError("This form of data is not supported")

        estimation = estimate_values(self.optresults, self.useful, data)

        self.estimation = estimation

        self.predictions = predict_classes(self.optresults, self.useful, estimation)

        return self.predictions

    def altsestimation(self):
        '''Returns the estimations of every alternative that has been calculated
        on the latest use of the predict method and contains an np array with the
        values'''

        return self.estimation


def criteria_plot(res):
    criterion_index = 0
    for criterion in range(len(res.g)):
        x_index = np.arange(len(res.g[criterion]))
        g_criterion = [ '%.3f' % elem for elem in res.g[criterion] ]

        criterion_title  = res.criteria[criterion]
        criterion_end_of_array = len(res.g[criterion])+criterion_index
        x_values = x_index
        y_values = res.crit_utilities[criterion_index : criterion_end_of_array]
        plt.plot(x_values,y_values)
        plt.xticks(x_index, g_criterion)
        plt.title(criterion_title)
        plt.show()

        criterion_index += len(res.g[criterion])

    return 0


def alternatives_plot(res):

    if hasattr(res, 'optresults'):
        x_indexes = np.arange(len(res.predictions))
        x_values = res.predictions
        y_values = res.estimation
        labels = res.predictions
        fig, ax = plt.subplots()
        for point in res.optresults.transitionpoints:
            ax.axhline(point, linestyle='--',color='gray', linewidth=1)
    else:
        x_indexes = np.arange(len(res.alternatives))
        x_values = res.alternatives
        y_values = res.alts_utilities
        labels = res.alternatives

    plt.bar(x_indexes, y_values)

    plt.xticks(ticks=x_indexes, labels=labels, rotation='vertical')
    plt.title('Alternatives')
    plt.show()

    return 0



