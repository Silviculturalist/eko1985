def jonsbon_to_decimal(jonsbon: float):
    n = 5 - jonsbon
    return 1.2**n


def jonson_1914_bonitet_mean_height_age_100(jonsbon: float):
    if jonsbon == 2:
        n = 27.7
    elif jonsbon == 3:
        n = 22.4
    elif jonsbon == 4:
        n = 18.0
    elif jonsbon == 5:
        n = 14.6
    elif jonsbon == 6:
        n = 11.6
    elif jonsbon == 7:
        n = 9.0
    elif jonsbon == 8:
        n = 7.0

    return n
