# -*- coding: utf-8 -*-
"""Module that testing creating of excel with default parameters."""
from loguru import logger
from create_excel import create_excel, Maslul, MortgageTypes


@logger.catch
def main():
    """function that running all tests"""
    create_excel('calc.xlsx',[Maslul(MortgageTypes.KALAZ,                      4,     20*12, 450000),
                              Maslul(MortgageTypes.PRAIM,                      1,     30*12, 250000),
                              Maslul(MortgageTypes.FIVE_YEARS_PERCENT,         3.5,   20*12, 100000),
                              Maslul(MortgageTypes.FIVE_YEARS_PERCENT_ZAMUD,   3,     15*12, 100000),
                              ])

if __name__ == "__main__":
    main()
