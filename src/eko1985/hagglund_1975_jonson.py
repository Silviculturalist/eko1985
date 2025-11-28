import numpy as np


def jonson_bonitet_pine_northern_Sweden(jonsbon: int,
                                        age: float) -> float:
    if age < 30 or age > 130:
        Warning("Age out of bounds for Jonsbon according to Hagglund 1975 reference materia (30-130)")
    a = 0
    if jonsbon == 3:
        a = 24.43
    elif jonsbon == 4:
        a = 21.45
    elif jonsbon == 5: 
        a = 18.64
    elif jonsbon == 6:
        a = 15.68
    elif jonsbon == 7:
        a = 12.69
    else:
        return ValueError("Invalid Jonsbon (3-7)")
    
    return a + 0.06506 * (age - 45 - abs(age-45)) - 0.01504*(age-45+abs(age-45))


def jonson_bonitet_pine_southern_Sweden(jonsbon: int,
                                        age : float) -> float:
    if age < 30 or age > 110:
        Warning("Age out of bounds for Jonsbon according to Hagglund 1975 reference material (30-110)")
    a = 0
    if jonsbon == 2:
        a = 26.82
    elif jonsbon == 3:
        a = 24.56
    elif jonsbon == 4:
        a = 21.78
    elif jonsbon == 5:
        a = 18.27
    elif jonsbon == 6:
        a = 15.87
    elif jonsbon == 7:
        a = 13.40
    else:
        return ValueError("Invalid Jonsbon (2-7)")
    
    return a - 0.02440*(age-45 + abs(age-45))

def jonson_bonitet_spruce_northern_Sweden(jonsbon:int,
                                          age : float) -> float:
    if age < 40 or age > 130:
        Warning("Age out of bounds for Jonsbon according to Hagglund 1975 reference material (40-130)")
    a = 0
    if jonsbon == 2:
        a = 28.69
    elif jonsbon == 3:
        a = 24.98
    elif jonsbon == 4:
        a = 21.87
    elif jonsbon == 5:
        a = 19.32
    elif jonsbon == 6:
        a = 16.79
    elif jonsbon == 7:
        a = 15.16
    else:
        return ValueError("Invalid Jonsbon (2-7)")
    
    return a - 0.02788*(age-60 + abs(age-60))
    

def jonson_bonitet_spruce_southern_Sweden(jonsbon:int,
                                          age : float) -> float:
    if age < 40 or age > 100:
        Warning("Age out of bounds for Jonsbon according to Hagglund 1975 reference material (40-100)")
    a = 0
    if jonsbon == 1:
        a = 33.40
    if jonsbon == 2:
        a = 29.56
    elif jonsbon == 3:
        a = 26.77
    elif jonsbon == 4:
        a = 23.64
    elif jonsbon == 5:
        a = 20.53
    else:
        return ValueError("Invalid Jonsbon (1-5)")
    
    return a + 0.05008 * ( age - 45 - abs(age-45)) - 0.03668*(age-45 + abs(age-45))


    
