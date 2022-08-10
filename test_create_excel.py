# -*- coding: utf-8 -*-
from loguru import logger
from create_excel import *


@logger.catch
def main():
    create_excel('calc.xlsx',[Maslul(MortgageTypes.Kalaz,                   4,   20*12, 450000),
                              Maslul(MortgageTypes.Praim,                   1,   30*12, 250000),
                              Maslul(MortgageTypes.FiveYearsPercent,      3.5,   20*12, 100000),
                              Maslul(MortgageTypes.FiveYearsPercentZamud,   3,   15*12, 100000),
                              ])
    pass
    
if __name__ == "__main__":
    main()