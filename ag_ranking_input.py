import skcriteria as skc
from skcriteria.preprocessing import invert_objectives, scalers
from skcriteria.madm import similarity
from skcriteria.pipeline import mkpipe

matrix = [[1,2,3],[4,5,6]]
print(matrix)
objectives = [min, max, min]
print(objectives)
weights = [0.5, 0.05, 0.45]
print(weights)
alternatives = ["car 0", "car 1"]
print(alternatives)
criteria = ["autonomy", "comfort", "price"]

dm = skc.mkdm(matrix, objectives, weights, alternatives, criteria)
print(dm)

pipe = mkpipe(
    invert_objectives.NegateMinimize(),
    scalers.VectorScaler(target="matrix"),  # this scaler transform the matrix
    scalers.SumScaler(target="weights"),  # and this transform the weights
    similarity.TOPSIS(),
)
print(pipe)

rank = pipe.evaluate(dm)
print(rank)