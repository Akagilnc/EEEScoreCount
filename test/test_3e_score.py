from source.eee_score import EEEScore
import pandas as pd
from os import listdir

test = EEEScore()


def test_main():
    result = test.main()
    assert result is not None
    print(result[[test.teachers_total_score_column_name, '日常分', '现场投票', test.total_score_column_name]])
    assert 'ranking.xlsx' in listdir('.')



def test_get_files():
    result = test.get_files('../files/')
    assert result
