from datetime import time

class gas_sampler():
    """
    Gas Sampler
    """
    def __init__(self, code_name: str, startup_time: time) -> None:
        self.code_name: str = code_name
        self.name: str = '大气采样器'
        self.sample_type: str = '大气'
        self.ports: list[int] = [1, 2]
        self.startup_time: time = startup_time
    
    def do_sample(self):
        pass

class dust_sampler():
    """
    Dust Sampler
    """
    def __init__(self, code_name: str, startup_time: time) -> None:
        self.code_name: str = code_name
        self.name: str = '粉尘采样器'
        self.sample_type: str = '粉尘'
        self.ports: list[int] = [1, 2]
        self.startup_time: time = startup_time
    pass

