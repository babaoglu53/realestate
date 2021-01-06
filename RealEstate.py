from abc import ABCMeta, abstractmethod

class RealEstate(metaclass = ABCMeta):

    @abstractmethod
    def getOffices(self) -> None:
        pass

    @abstractmethod
    def getAgents(self) -> None:
        pass