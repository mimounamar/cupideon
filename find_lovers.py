from gale_shapley import start
from letters import generate_letters
from retrieve_data import retrieve_data
from score_calculator import calculate

retrieve_data()
results = calculate()
matches = start(results["girls"], results["boys"], results["attribution"])
generate_letters(matches)