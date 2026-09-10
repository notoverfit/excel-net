import pandas as pd
from sklearn.datasets import make_circles

X, y = make_circles(n_samples=10000, factor=0.5, noise=0.08, random_state=10)

concentric_circles = pd.DataFrame({"x1": X[:, 0], "x2": X[:, 1], "y": y})
concentric_circles.to_csv('./data-01.csv', index=False)