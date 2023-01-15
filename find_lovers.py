from gale_shapley import start
from score_calculator import calculate

results = calculate()
start(results["girls"], results["boys"], results["attribution"])