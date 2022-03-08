from typing import List
from dataclasses import dataclass
import re


@dataclass
class Bid:
    company: str
    counter_party: str
    price: str
    price_type: str
    delivery_window: str
    quantity: str
    discharge_port: str

    @staticmethod
    def parse_string(bid_str):
        REX = [
            re.compile(r"DELIVERY WINDOW: (?P<del_window>.*\d{4})"),
            re.compile(r"DISCHARGE PORT: (?P<discharge_port>.*) LOAD PORT"),
            re.compile(r"LOAD PORT: (?P<load_port>.*) QUANTITY"),
            re.compile(r"QUANTITY: (?P<quantity>.*tolerance)")
        ]
        #(raises|lowers) bid (?P<DelStart>[A-Za-z]{3} \d\d?)-(?P<DelEnd>[A-Za-z]{3} \d\d?)
        #https://jsonlint.com/
        #https://regex101.com/
        JKM_PRICE_TYPE_REX = re.compile(r"JKM Full Mnth (?P<month>.{3})")

        def clean_paren(s):
            if "(" in s:
                return s.split("(")[0]
            else:
                return s

        def get_company():
            return bid_str.split(",")[1].strip().split(" ")[0]

        def get_price_type():
            if "Flat Price" in bid_str:
                return "Flat Price"
            else:
                matches = re.finditer(JKM_PRICE_TYPE_REX, bid_str)
                for m in matches:
                    month = m.groupdict()["month"]
                    return f"JKM {month}"

        def get_price():
            return bid_str.split("$")[1].strip().split(" ")[0]

        found = {}
        for pattern in REX:
            matches = re.finditer(pattern, bid_str)
            for match in matches:
                found.update(match.groupdict())

        found = {k: clean_paren(v).strip() for k, v in found.items()}
        return Bid(
            company=get_company(),
            counter_party="",
            price=get_price(),
            price_type=get_price_type(),
            delivery_window=found.get("del_window", ""),
            quantity=found.get("quantity", ""),
            discharge_port=found.get("discharge_port", ""),
        )


@dataclass
class MOCEmailNotificationBody:
    bids: List[str]
    offers: List[str]
    trades: List[str]
    withdrawal: List[str]
    exclusion: List[str]
    @staticmethod
    def parse_body(body):
        bids = []
        offers = []
        exclusion = []
        withdrawal = []
        trades = []
        current = None
        for line in body.split("\n"):
            if "BIDS" in line:
                current = "BIDS"
            elif "OFFERS" in line:
                current = "OFFERS"
            elif "TRADES" in line:
                current = "TRADES"
            elif "WITHDRAWALS" in line:
                current = "WITHDRAWALS"
            elif "EXCLUSIONS" in line:
                current = "EXCLUSIONS"
            elif "Derivative" in line:
                break
            elif current is not None and len(line.strip()) > 1:  # ignore \r (carriage return)
                line = line.strip()
                if current == "BIDS":
                    bids.append(line)
                elif current == "OFFERS":
                    offers.append(line)
                elif current == "TRADES":
                    trades.append(line)
                elif current == "WITHDRAWALS":
                    withdrawal.append(line)
                elif current == "EXCLUSIONS":
                    exclusion.append(line)

        return MOCEmailNotificationBody(bids, offers, trades,withdrawal,exclusion)
