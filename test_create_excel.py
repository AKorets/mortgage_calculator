# -*- coding: utf-8 -*-
"""Module that testing creating of excel with default parameters."""
from loguru import logger
from create_excel import create_excel, Maslul, MortgageTypes


@logger.catch
def main():
    """function that running all tests"""
    create_excel('calc.xlsx',[Maslul(MortgageTypes.Kalaz,                   4,   20*12, 450000),
                              Maslul(MortgageTypes.Praim,                   1,   30*12, 250000),
                              Maslul(MortgageTypes.FiveYearsPercent,      3.5,   20*12, 100000),
                              Maslul(MortgageTypes.FiveYearsPercentZamud,   3,   15*12, 100000),
                              ])

if __name__ == "__main__":
    main()
