from KwTurkiye import KwTurkiye
from Era import Era
from Cb import Cb
from Premar import Premar
from KristalTurkiye import KristalTurkiye
from Turyap import Turyap
from TamNokta import TamNokta
from RealtyWorld import RealtyWorld
from CenturyGlobal import CenturyGlobal

class RealEstateFactory:

    def createRealEstateObject(self, name):
        if name == 'kwturkiye':
            return KwTurkiye()

        if name == 'era':
            return Era()

        if name == 'cb':
            return Cb()

        if name == 'premar':
            return Premar()

        if name == 'kristalturkiye':
            return KristalTurkiye()

        if name == 'turyap':
            return Turyap()

        if name == 'tamnokta':
            return TamNokta()

        if name == 'realtyworld':
            return RealtyWorld()

        if name == 'centuryglobal':
            return CenturyGlobal()