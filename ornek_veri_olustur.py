"""Demo için gerçekçi ama dağınık şube satış dosyaları üretir.

Gerçek hayattaki sorunları bilerek ekler: farklı sütun adları, farklı tarih
formatları, boş satırlar, tekrar eden kayıtlar, yazım farkları.
"""
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)
GIRDI = Path(__file__).parent / "girdi"

URUNLER = {
    "Çay 1kg": 185.0, "Kahve 250g": 240.0, "Şeker 5kg": 210.0,
    "Un 10kg": 320.0, "Zeytinyağı 1L": 450.0, "Makarna 500g": 28.5,
    "Pirinç 2kg": 115.0, "Peynir 1kg": 390.0,
}

# Her şube kendi formatında gönderiyor
SUBELER = {
    "Merkez":  {"sutunlar": ["Tarih", "Ürün", "Adet", "Birim Fiyat"], "tarih": "%d.%m.%Y"},
    "Safranbolu": {"sutunlar": ["tarih", "urun_adi", "miktar", "fiyat"], "tarih": "%Y-%m-%d"},
    "Eskipazar": {"sutunlar": ["TARİH", "ÜRÜN", "ADET", "FİYAT (TL)"], "tarih": "%d/%m/%Y"},
}


def sube_dosyasi(sube: str, ayar: dict, gun_sayisi: int = 30) -> pd.DataFrame:
    baslangic = date(2026, 9, 1)
    satirlar = []
    for g in range(gun_sayisi):
        gun = baslangic + timedelta(days=g)
        for _ in range(random.randint(3, 9)):
            urun = random.choice(list(URUNLER))
            # Yazım farkları: küçük harf, fazladan boşluk
            yazim = random.choice([urun, urun.lower(), f" {urun} "])
            satirlar.append([
                gun.strftime(ayar["tarih"]), yazim,
                random.randint(1, 20), URUNLER[urun],
            ])
    df = pd.DataFrame(satirlar, columns=ayar["sutunlar"])
    # Tekrarlanan kayıtlar ve boş satırlar
    df = pd.concat([df, df.sample(5, random_state=1)], ignore_index=True)
    # Araya serpiştirilmiş boş satırlar
    for konum in random.sample(range(1, len(df)), 3):
        df = pd.concat([df.iloc[:konum], pd.DataFrame([[None] * 4], columns=df.columns), df.iloc[konum:]])
    return df.reset_index(drop=True)


if __name__ == "__main__":
    GIRDI.mkdir(exist_ok=True)
    for sube, ayar in SUBELER.items():
        yol = GIRDI / f"{sube}_eylul.xlsx"
        sube_dosyasi(sube, ayar).to_excel(yol, index=False)
        print(f"oluşturuldu: {yol.name}")
