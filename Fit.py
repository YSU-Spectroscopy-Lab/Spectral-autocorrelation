import numpy as np
from scipy.optimize import leastsq

def linear_func(params, x):
    k, b = params
    return k * x + b

def error_func(params, x, y):
    return linear_func(params, x) - y

initial_params = [1, 1]

x = np.array([2, 3, 4, 5, 6])
y = np.array([2966.20, 3862.28, 4967.44, 5898.85, 6921.60])  # y坐标

params_fit, success = leastsq(error_func, initial_params, args=(x, y))

k_fit, b_fit = params_fit

print("k_fit, b_fit",k_fit, b_fit)