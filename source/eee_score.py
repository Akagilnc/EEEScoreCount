import pandas as pd
import numpy as np
from os import listdir


class EEEScore:
    DEFAULT_COLUMN = ['扩展属性', '姓名', '创业背景描述', '创业项目描述', '团队介绍', '需求描述', '现场情况', '现场情况.1', '现场情况.2']
    DEFAULT_DAILY_COLUMN = ['姓名', '日常分数', '线上分数', '日常分', '现场投票']
    path = '../files/'
    teacher_data = None
    daily_data = None
    teachers_total_score_column_name = '导师评分'
    total_score_column_name = 'total_score'
    daily_score_column_name = '日常分'
    voting_score_column_name = '现场投票'

    def main(self):
        files = self.get_files()
        for file in files:
            if '~$' == file[:2]:
                continue
            if '毕业典礼打分表' in file and '第七期' in file:
                self.teacher_data = self.get_teachers_df(self.read_excel(file))
            if '训练营' in file and '11月评分' in file:
                self.daily_data = self.get_daily_score(self.read_excel(file, '日常分数+线上分数排名', 1))
        if not self.teacher_data or self.teacher_data.empty:
            raise ValueError('No files found')
        result = self.get_teachers_total_score(self.teacher_data)
        result = self.get_final_teacher_score(result)
        result = result.join(self.daily_data)
        result = self.get_total_score(result)
        result = self.sort_data(result)
        self.make_ranking_file(result)
        return result

    def read_excel(self, file_name, sheet_name='Sheet1', header=0):
        file_path = self.path + file_name
        return pd.read_excel(file_path, sheet_name=sheet_name, header=header)

    def get_teachers_df(self, df, columns=None):
        if columns is None:
            columns = self.DEFAULT_COLUMN
        return df[columns]

    def get_teachers_total_score(self, df):
        weights = np.array([[1, 3, 1.5, 1.5, 1, 1, 1]]).T
        scores = df[self.DEFAULT_COLUMN[2:]]
        teachers_total_scores = np.dot(scores, weights)

        df[self.teachers_total_score_column_name] = teachers_total_scores * 0.65
        return df

    def get_total_score(self, df):
        scores = df[[self.teachers_total_score_column_name,
                     self.daily_score_column_name,
                     self.voting_score_column_name]]
        df[self.total_score_column_name] = scores.sum(axis=1)
        return df

    def get_voting_score(self, df):
        voting_score = df[self.voting_score_column_name]
        return voting_score / voting_score.max() * 10

    def sort_data(self, df):
        return df.sort_values(by=self.total_score_column_name, ascending=False)

    def get_files(self, file_path=None):
        return listdir(file_path if file_path else self.path)

    @staticmethod
    def get_final_teacher_score(df):
        return df.groupby(by="姓名").mean()

    def get_daily_score(self, df):
        df = df[self.DEFAULT_DAILY_COLUMN].set_index('姓名')
        df[self.voting_score_column_name] = self.get_voting_score(df)
        return df

    @staticmethod
    def make_ranking_file(df):
        df = df.reset_index()
        df.index = df.index + 1
        df.to_excel('ranking.xlsx')

