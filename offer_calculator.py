from dataclasses import dataclass
from typing import Dict, Optional
import pandas as pd


@dataclass
class BaseRates:
    """Базовые цены за единицу ресурса"""
    gems: float = 0.0125
    skip: float = 0.3
    tnt: float = 0.06
    nitro: float = 0.03
    no_ads: float = 10.0
    small_chest: float = 0.75  # 60 gems
    big_chest: float = 3.75    # 300 gems
    no_chest: float = 0.0


@dataclass
class OfferResources:
    """Ресурсы в офере"""
    gems: int = 0
    skip: int = 0
    tnt: int = 0
    nitro: int = 0
    no_ads: int = 0
    chest_type: str = "NoChest"  # NoChest, Small, Big
    chest_amount: int = 0


class OfferCalculator:
    def __init__(self, base_rates: Optional[BaseRates] = None):
        self.base_rates = base_rates or BaseRates()

    def calculate_base_value(self, resources: OfferResources) -> float:
        """Рассчитывает базовую стоимость всех ресурсов"""
        value = 0.0

        value += resources.gems * self.base_rates.gems
        value += resources.skip * self.base_rates.skip
        value += resources.tnt * self.base_rates.tnt
        value += resources.nitro * self.base_rates.nitro
        value += resources.no_ads * self.base_rates.no_ads

        chest_rates = {
            "Small": self.base_rates.small_chest,
            "Big": self.base_rates.big_chest,
            "NoChest": self.base_rates.no_chest
        }
        value += resources.chest_amount * chest_rates.get(resources.chest_type, 0)

        return round(value, 2)

    def calculate_discount(self, pack_price: float, base_value: float) -> float:
        """Рассчитывает скидку в процентах"""
        if base_value == 0:
            return 0.0
        discount = ((base_value - pack_price) / base_value) * 100
        return round(discount, 0)

    def calculate_roi(self, pack_price: float, base_value: float) -> float:
        """Рассчитывает ROI (Return on Investment)"""
        if pack_price == 0:
            return 0.0
        roi = ((base_value - pack_price) / pack_price) * 100
        return round(roi, 0)

    def get_multiplier(self, roi: float) -> str:
        """Определяет множитель на основе ROI"""
        if roi >= 1500:
            return "17x VALUE"
        elif roi >= 1400:
            return "15x VALUE"
        elif roi >= 700:
            return "8x VALUE"
        elif roi >= 600:
            return "6x VALUE"
        elif roi >= 400:
            return "5x VALUE"
        elif roi >= 300:
            return "4x VALUE"
        elif roi >= 200:
            return "3x VALUE"
        elif roi >= 150:
            return "2x VALUE"
        else:
            return "1x VALUE"

    def get_value_badge(self, roi: float) -> str:
        """Создаёт бейдж с процентами"""
        return f"{int(roi)}% MORE!"

    def determine_pack_type(self, resources: OfferResources) -> str:
        """Определяет тип пака на основе преобладающего ресурса"""
        if resources.skip >= 20 and resources.skip > resources.gems * 0.03:
            return "SkipIt Focus"
        elif resources.gems >= 3000:
            return "Gem Focus"
        elif resources.tnt > 0 and resources.nitro > 0:
            return "Balanced"
        elif resources.tnt > resources.nitro * 2:
            return "TNT Focus"
        elif resources.nitro > resources.tnt * 2:
            return "NITRO Focus"
        else:
            return "Balanced"

    def check_min_skipits(self, skip_amount: int) -> str:
        """Проверяет минимальное количество SkipIts"""
        return "✓" if skip_amount >= 20 else "✖"

    def calculate_offer(self, name: str, resources: OfferResources, pack_price: float) -> Dict:
        """Полный расчёт оффера"""
        base_value = self.calculate_base_value(resources)
        discount = self.calculate_discount(pack_price, base_value)
        roi = self.calculate_roi(pack_price, base_value)

        return {
            "Name": name,
            "Gems": resources.gems,
            "Skip": resources.skip,
            "TNT": resources.tnt,
            "Nitro": resources.nitro,
            "No Ads": resources.no_ads,
            "Chest Type": resources.chest_type,
            "Chest Amount": resources.chest_amount,
            "Pack Price ($)": pack_price,
            "Base Value ($)": base_value,
            "Discount (%)": discount,
            "ROI (%)": roi,
            "Value Badge": self.get_value_badge(roi),
            "Multiplier": self.get_multiplier(roi),
            "Min SkipIts": self.check_min_skipits(resources.skip),
            "Pack Type": self.determine_pack_type(resources)
        }


def load_existing_offers(file_path: str) -> pd.DataFrame:
    """Загружает существующие офферы из Excel файла"""
    df = pd.read_excel(file_path, sheet_name='O.Calculator', header=None)

    offers = []
    current_section = None

    for idx, row in df.iterrows():
        if idx == 0 or pd.isna(row[0]):
            continue

        if "STARTER PACK" in str(row[0]) or "CHEST + SKIPIT" in str(row[0]) or "BLITZ OFFER" in str(row[0]) or "CUSTOM" in str(row[0]):
            current_section = str(row[0])
            continue

        if current_section and not pd.isna(row[3]):
            try:
                pack_price = float(row[3])
                if pack_price > 0:
                    offer = {
                        "Section": current_section,
                        "Name": str(row[0]),
                        "Resource 1": row[1],
                        "Resource 2": row[2],
                        "Pack Price ($)": pack_price,
                        "Base Value ($)": row[4] if not pd.isna(row[4]) else 0,
                        "Discount": row[5] if not pd.isna(row[5]) else 0,
                        "Value Badge": row[6] if not pd.isna(row[6]) else "",
                        "Multiplier": row[7] if not pd.isna(row[7]) else "",
                        "Extra": row[8] if not pd.isna(row[8]) else "",
                        "Pack Type": row[9] if not pd.isna(row[9]) else ""
                    }
                    offers.append(offer)
            except (ValueError, TypeError):
                continue

    return pd.DataFrame(offers)
