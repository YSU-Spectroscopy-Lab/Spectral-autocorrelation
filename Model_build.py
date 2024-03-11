from sklearn.model_selection import train_test_split
import pandas as pd
import numpy as np
from keras.models import Model
from keras.utils import plot_model
from keras.layers import Dense, Flatten, Conv1D, MaxPooling1D, Input


def load_total_data(path, num):
    df = pd.read_pickle(path)
    x = np.expand_dims(df.values[:, 0:-1].astype(float), axis=2)  # Adding a one-dimensional axis
    y = df.values[:, -1] / num
    # Divide training set, test set
    x_train, x_test, y_train, y_test = train_test_split(x, y, test_size=0.05, shuffle=True)
    print("Loading of data complete!")
    return x_train, x_test, y_train, y_test


def load_total_data_multi(path, num):
    df = pd.read_pickle(path)
    x = np.expand_dims(df.values[:, 0:-2].astype(float), axis=2)  # Adding a one-dimensional axis
    y = df.values[:, -2:] / num
    # Divide training set, test set
    x_train, x_test, y_train, y_test = train_test_split(x, y, test_size=0.05, shuffle=True)
    print("Loading of data complete!")
    return x_train, x_test, y_train, y_test


def build_CNN_model(model_structure):
    input1 = Input(shape=(685, 1))
    conv_layer1_1 = Conv1D(16, 3, strides=1, activation='relu')(input1)
    # conv_layer1_1_ = Conv1D(16, 3, strides=1, activation='relu')(conv_layer1_1)
    max_layer1_1 = MaxPooling1D(3)(conv_layer1_1)
    conv_layer1_2 = Conv1D(32, 3, strides=1, activation='relu')(max_layer1_1)
    # conv_layer1_2_ = Conv1D(32, 3, strides=1, activation='relu')(conv_layer1_2)
    max_layer1_2 = MaxPooling1D(3)(conv_layer1_2)
    conv_layer1_3 = Conv1D(32, 3, activation='relu')(max_layer1_2)
    # conv_layer1_3_ = Conv1D(32, 3, activation='relu')(conv_layer1_3)
    max_layer1_3 = MaxPooling1D(3)(conv_layer1_3)
    conv_layer1_4 = Conv1D(32, 3, activation='relu')(max_layer1_3)
    # conv_layer1_4_ = Conv1D(32, 3, activation='relu')(conv_layer1_4)
    max_layer1_4 = MaxPooling1D(3)(conv_layer1_4)
    flatten = Flatten()(max_layer1_4)
    f1 = Dense(1, activation='linear', name='prediction_one')(flatten)
    model = Model(outputs=f1, inputs=input1)
    model.summary()
    plot_model(model, to_file=model_structure, show_shapes=True)  # Printed model structure
    return model


def build_CNN_model_multi(model_structure):
    input1 = Input(shape=(685, 1))
    conv_layer1_1 = Conv1D(16, 3, strides=1, activation='relu')(input1)
    # conv_layer1_1_ = Conv1D(16, 3, strides=1, activation='relu')(conv_layer1_1)
    max_layer1_1 = MaxPooling1D(3)(conv_layer1_1)
    conv_layer1_2 = Conv1D(32, 3, strides=1, activation='relu')(max_layer1_1)
    # conv_layer1_2_ = Conv1D(32, 3, strides=1, activation='relu')(conv_layer1_2)
    max_layer1_2 = MaxPooling1D(3)(conv_layer1_2)
    conv_layer1_3 = Conv1D(32, 3, activation='relu')(max_layer1_2)
    # conv_layer1_3_ = Conv1D(32, 3, activation='relu')(conv_layer1_3)
    max_layer1_3 = MaxPooling1D(3)(conv_layer1_3)
    conv_layer1_4 = Conv1D(32, 3, activation='relu')(max_layer1_3)
    # conv_layer1_4_ = Conv1D(32, 3, activation='relu')(conv_layer1_4)
    max_layer1_4 = MaxPooling1D(3)(conv_layer1_4)
    flatten = Flatten()(max_layer1_4)
    f1 = Dense(2, activation='linear', name='prediction_one')(flatten)
    model = Model(outputs=f1, inputs=input1)
    model.summary()
    plot_model(model, to_file=model_structure, show_shapes=True)  # Printed model structure
    return model


def run(train_no_model, train_nh3_model, train_no_nh3_model, model_structure):
    if train_no_model:
        data_no_pata = 'Pkl_single_component_data/no-spectrum.pkl'
        x_train, x_test, y_train, y_test = load_total_data(data_no_pata, 1000)
        model = build_CNN_model(model_structure)

        return model, x_train, x_test, y_train, y_test, 1000

    if train_nh3_model:
        data_nh3_pata = 'Pkl_single_component_data/nh3-spectrum.pkl'
        x_train, x_test, y_train, y_test = load_total_data(data_nh3_pata, 10000)
        model = build_CNN_model(model_structure)

        return model, x_train, x_test, y_train, y_test, 10000

    if train_no_nh3_model:
        data_no_nh3_pata = 'Pkl_multicomponent_data/no-nh3-spectrum.pkl'
        x_train, x_test, y_train, y_test = load_total_data_multi(data_no_nh3_pata, 10000)
        model = build_CNN_model_multi(model_structure)

        return model, x_train, x_test, y_train, y_test, 10000
