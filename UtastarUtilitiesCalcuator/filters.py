import pandas as pd
import numpy as np
from pandas.api.types import is_string_dtype
from pandas.api.types import is_numeric_dtype


def swap(l, data, pos1, pos2):
    l[pos1], l[pos2] = l[pos2], l[pos1]
    data[[pos1, pos2]] = data[[pos2, pos1]]
    return l, data


def sort(l, data):
    for i in range(len(l)):
        k = i
        for j in range(i+1, len(l)):
            if l[k] > l[j]:
                k = j
        if k != i:
            l, data = swap(l, data, k, i)
    return l, data


class CheckData(object):

    def __init__(self, data):
        self.data = data

    def ranking_pd(self):
        con_ranking = 'Ranking' in self.data.columns
        con_rank = 'Rank' in self.data.columns
        if con_rank:
            self.data.rename(columns={'Rank': 'Ranking'}, inplace=True)
            con_ranking = True
        if con_ranking:
            max_value = self.data["Ranking"].max()
            for expected_value in range(1, max_value):
                arg = not (expected_value in self.data["Ranking"].values)
                if arg:
                    ValueError("There are missing values in the ranking column")

            self.data = self.data.sort_values("Ranking")
            return self.data

        raise ValueError("There is not an acceptable ranking column with labels Rank or Ranking")

    def reposition_ranking_pd(self):

        last = self.data['Ranking']
        self.data.drop(['Ranking'], axis=1, inplace=True)
        position = len(self.data.columns)
        self.data.insert(position, 'Ranking', last)
        return self.data

    def filter_values_pd(self):
        ### 2. Empty places check  and character check  ###

        # Check NaN values
        if self.data.isnull().values.any():
            raise ValueError("There are null values ")
        # check characters values, mixed character and numeric values
        char_columns = 0
        b = is_numeric_dtype(self.data.index)
        d = is_string_dtype(self.data.index)
        for col in self.data.columns:
            if is_string_dtype(self.data[col]):
                if char_columns == 0:
                    alt = col
                char_columns += 1
                a = char_columns > 1
                c = char_columns > 0
                if a and b:
                    raise ValueError("There ara more than one none numerical columns")
                if c and d:
                    raise ValueError("There must not be none numerical columns since the index is characters")

         # Create a column of alternatives if it does not exist and move
         # the alternatives column to first place if it exists
        if char_columns == 0 and b:
            column_num = len(self.data)
            alts_list = ['alternative' + str(i) for i in range(1, column_num + 1)]
            alts_pandas = pd.DataFrame(alts_list, columns=["alternatives"])
            self.data.insert(0, 'alternatives', alts_pandas)

        if char_columns == 1:
            first = self.data[alt]
            self.data.drop([alt], axis=1, inplace=True)
            self.data.insert(0, alt, first)

        return self.data

    def cut_the_ranking(self):
        ranking = self.data.iloc[:, -1].tolist()
        self.data = self.data.drop(self.data.columns[-1], axis=1)
        return ranking

    def turn_to_np(self):
        ###  data transformation from pandas to numpy  ###
        if is_numeric_dtype(self.data.index):
            alts_list = self.data.iloc[:, 0].tolist()
            self.data = self.data.drop(self.data.columns[0], axis=1)
            criteria = list(self.data.columns.values)
            self.data = self.data.to_numpy()
            return self.data, criteria, alts_list

        if is_string_dtype(self.data.index):
            alts_list = self.data.index.tolist()
            self.data = self.data
            criteria = list(self.data.columns.values)
            self.data = self.data.to_numpy()
            return self.data, criteria, alts_list

        raise ValueError('Index of dataframe has unknown data type')

    # Filter for the alternatives and the criteria lists

    def list_filter(self, list_of_things, dimension, key_word):
        list_size = np.size(self.data, dimension)
        if list_of_things is None:
            list_of_things = [key_word + str(i) for i in range(1, list_size + 1)]
            return list_of_things

        if list_size != len(list_of_things):
            raise ValueError("The {} list size does not much the size of the array".format(key_word))

        return list_of_things

    # Ranking check for the np arrays
    def ranking_np(self, list_of_things):
        list_size = np.size(self.data, 0)
        if list_of_things is None:
            list_of_things = [i for i in range(1, list_size+1)]
            return list_of_things

        if list_size != len(list_of_things):
            raise ValueError("The ranking size does not much the multi-criteria array size")

        maxi = int(max(list_of_things))
        for i in range(1, maxi+1):
            if i not in list_of_things:
                raise ValueError("The ranking list is missing the value {}".format(i))
        return list_of_things

    def filter_values_np(self, ranking):
        '''
        This method checks the numpy array for nan values and returns back
         tha array and the ranking list sorted

        :param ranking: Accepts the ranking of the alternatives as a list
        :return: The data and the ranking sorted
        '''
        sum = np.sum(self.data)
        if np.isnan(sum):
            raise ValueError("There are Nan values in the array ")
        else:
            ranking, self.data = sort(ranking, self.data)

            return ranking, self.data

    def gradient(self, monotonicity):

        '''  This method controls the values of the list  monotonicity, the list accepts as
            values either 1 for criteria that their values become beneficial when maximized
            and 0 for criteria that their values become beneficial when minimized '''

        length = np.size(self.data, 1)
        if monotonicity is None:
            monotonicity = [1 for i in range(length)]
            return monotonicity
        else:
            if length != len(monotonicity):
                raise ValueError("The size of monotonicity does not match the size of the criteria data array ")
            for i in monotonicity:
                if i != 0 and i != 1:
                    raise ValueError("There are values that are not 1 nor 0 in monotonicity list")

            return monotonicity

    def BestValues(self, bestvalues, monotonicity):

        ''' This method checks if the values of the list bestvalues is None or not. If None it creates a list
         of the best values depending on the monotonicity. If the list already contains values it checks the
          size of the list and the values it contains '''

        length = np.size(self.data, 1)
        if bestvalues is None:
            bestvalues = []
            for column in range(0, length):
                if monotonicity[column] == 1:
                    bestvalues.append(np.amax(self.data[:, column]))
                elif monotonicity[column] == 0:
                    bestvalues.append(np.amin(self.data[:, column]))

            return bestvalues

        else:
            if len(bestvalues) is not length:
                raise ValueError("The size of the bestvalues does not much with the number of criteria")

            else:
                for column in range(0, length):
                    if monotonicity[column] == 1 and bestvalues[column] < np.amax(self.data[:, column]):
                        raise ValueError("There are values out of bound in column {} of the multi-criteria array".format(column))
                    if monotonicity[column] == 0 and bestvalues[column] > np.amin(self.data[:, column]):
                        raise ValueError("There are values out of bound in column {} of the multi-criteria array".format(column))
                return bestvalues

    def WorstValues(self, worstvalues, monotonicity):

        ''' This method checks if the values of the list worstvalues is None or not. If None it creates a list
         of the worst values depending on the monotonicity. If the list already contains values it checks the
          size of the list and the values it contains '''

        length = np.size(self.data, 1)
        if worstvalues is None:
            worstvalues = []
            for column in range(0, length):
                if monotonicity[column] == 1:
                    worstvalues.append(np.amin(self.data[:, column]))
                elif monotonicity[column] == 0:
                    worstvalues.append(np.amax(self.data[:, column]))

            return worstvalues

        else:
            if len(worstvalues) is not length:
                raise ValueError("The size of the worstvalues does not much with the number of criteria")
            else:
                for column in range(0, length):
                    if monotonicity[column] == 1 and worstvalues[column] > np.amin(self.data[:, column]):
                        raise ValueError("There are values out of bound in column {} of the multi-criteria array".format(column))
                    if monotonicity[column] == 0 and worstvalues[column] < np.amax(self.data[:, column]):
                        raise ValueError("There are values out of bound in column {} of the multi-criteria array".format(column))
            return worstvalues

    def Intervals(self, intervals):
        if intervals is None:
            intervals = [3 for i in range(self.data.shape[1])]
            return intervals
        else:
            return intervals

    def group_consistency(self,groups, pref_order):
        for group in groups:
            if group not in pref_order:
                raise ValueError("The groups list contains not defined values. Check pref_order for missing values")


    def filter_nan_UTADIS(self):
        # Check Nan values for UTADIS
        sum_of_values = np.sum(self.data)
        if np.isnan(sum_of_values):
            raise ValueError("There are Nan values in the array ")
        return self.data

    def bounds_UTADIS(self, g, monotonicity):
        '''beltiwsh twn sxoliwn'''
        for i in range(len(g)):
            if monotonicity[i] == 1:
                if g[i][-1] < np.amax(self.data[:, i]) or g[i][0] > np.amin(self.data[:, i]):
                    raise ValueError("There are values at column  that are out of bounds")
            elif monotonicity[i] == 0:
                if g[i][0] < np.amax(self.data[:, i]) or g[i][-1] > np.amin(self.data[:, i]):
                    raise ValueError("There are values at column  that are out of bounds")
        return self.data


class ObjectDictionary(object):

    def __init__(self, dictionary):
        for key in dictionary:
            setattr(self, key, dictionary[key])


class OptimizeResult(dict):
    """ Represents the optimization result.
    Attributes

 """

    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError:
            raise AttributeError(name)

    __setattr__ = dict.__setitem__
    __delattr__ = dict.__delitem__

    def __repr__(self):
        if self.keys():
            m = max(map(len, list(self.keys()))) + 1
            return '\n'.join([k.rjust(m) + ': ' + repr(v)
                              for k, v in self.items()])
        else:
            return self.__class__.__name__ + "()"

    def __dir__(self):
        return list(self.keys())
            







        









