from abc import ABC, abstractmethod


class CYAdapter(ABC):
    @abstractmethod
    def get_buy_price(self, cert_number: str) -> dict:
        raise NotImplementedError

    @abstractmethod
    def submit_cert_lookup(self, cert_number: str, slab_type: str) -> dict:
        raise NotImplementedError
