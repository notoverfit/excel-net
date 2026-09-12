import pandas as pd
import numpy as np
from pyprojroot import here
from sklearn.model_selection import train_test_split
from sklearn.datasets import make_circles

X, y = make_circles(n_samples=10000, factor=0.5, noise=0.08, random_state=10)

concentric_circles = pd.DataFrame({"x1": X[:, 0], "x2": X[:, 1], "y": y})
train, test = train_test_split(concentric_circles, test_size=0.2, random_state=10)
test.to_csv(here('data/circles-test.csv'), index=False)
train.to_csv(here('data/circles-train.csv'), index=False)

def make_spirals(n_samples=10000, noise=0.08, random_state=10):
    rng = np.random.default_rng(random_state)

    n = n_samples // 2

    theta = np.linspace(0, 4 * np.pi, n)
    radius = np.linspace(0.1, 1.0, n)

    x1 = radius * np.cos(theta)
    y1 = radius * np.sin(theta)

    x2 = radius * np.cos(theta + np.pi)
    y2 = radius * np.sin(theta + np.pi)

    X = np.vstack([
        np.column_stack([x1, y1]),
        np.column_stack([x2, y2])
    ])

    y = np.concatenate([
        np.zeros(n),
        np.ones(n)
    ])

    # Add Gaussian noise
    X += rng.normal(0, noise, X.shape)

    # Shuffle
    indices = rng.permutation(len(X))

    return X[indices], y[indices]


X, y = make_spirals(n_samples=10000, noise=0.08, random_state=10)

spirals = pd.DataFrame({"x1": X[:, 0], "x2": X[:, 1], "y": y.astype(int)})

train, test = train_test_split(spirals, test_size=0.2, random_state=10, stratify=spirals["y"])

test.to_csv(here("data/spirals-test.csv"), index=False)
train.to_csv(here("data/spirals-train.csv"), index=False)