from sklearn.datasets import load_iris


def load_iris_dataset():
    iris = load_iris(as_frame=True)
    return iris.frame, iris.target
