import unittest
import pandas as pd

from crm_data import (
    classify_industry,
    classify_industry_series,
    pick_client,
    pick_client_series,
    pick_industry,
    pick_industry_series,
)


class TestVectorizedCrmPreparation(unittest.TestCase):
    def setUp(self):
        self.frame = pd.DataFrame(
            {
                "Клиент РА": ["ПАО Газпром", None, "  ", "Агентство", None, None],
                "Клиент": ["fallback", "Клиент Б", "Клиент В", "fallback", None, None],
                "Рекламодатель": ["Рекл. А", "Рекл. Б", "Рекл. В", None, None, None],
                "Отрасли КлиентаРА": [
                    "Банк",
                    "Строительство",
                    "Реклама, маркетинг, выставки",
                    "Реклама",
                    None,
                    "Неизвестная отрасль",
                ],
                "Отрасли Клиента": ["Выставки", None, None, None, None, None],
                "Отрасли Рекламодателя": [None, None, None, None, None, None],
                "Позиция заказа": [
                    "Форум и стенд",
                    "Баннерное размещение",
                    "Обычная интеграция",
                    "EXPO event",
                    "Баннер",
                    "Публикация",
                ],
                "Описание": ["", "", "", "", "", "SMM"],
                "Номенклатура": ["", "", "", "", "", ""],
            },
            index=[101, 205, 309, 412, 518, 623],
        )

    def test_client_selection_preserves_priority_and_fallback(self):
        expected = self.frame.apply(pick_client, axis=1)
        actual = pick_client_series(self.frame)
        pd.testing.assert_series_equal(
            expected.astype(str), actual.astype(str), check_names=False
        )

    def test_industry_selection_preserves_priority_and_fallback(self):
        expected = self.frame.apply(pick_industry, axis=1)
        actual = pick_industry_series(self.frame)
        pd.testing.assert_series_equal(
            expected.astype(str), actual.astype(str), check_names=False
        )

    def test_vectorized_classifier_matches_legacy_row_classifier(self):
        legacy_frame = self.frame.copy()
        legacy_frame["ОТРАСЛЬ_КЛИЕНТА"] = pick_industry_series(legacy_frame)
        expected = legacy_frame.apply(classify_industry, axis=1)
        actual = classify_industry_series(legacy_frame)
        pd.testing.assert_series_equal(
            expected.astype(str), actual.astype(str), check_names=False
        )


if __name__ == "__main__":
    unittest.main()
