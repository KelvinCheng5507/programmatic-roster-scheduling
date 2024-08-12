from pyomo.environ import *
import pandas as pd
import xlsxwriter

# Your parameters
persons = ["John", "Peter", "Mary", "Josh"]
dates = pd.date_range('2024-01-01', '2024-01-31')

# Annual Leaves
AL = {
	'John': ['2024-01-01', '2024-01-02', '2024-01-03'],
	"Peter": ['2024-01-04',  '2024-01-11'],
	"Mary": ['2024-01-28', '2024-01-29', '2024-01-30'],
	"Josh": ['2024-01-15']
}

# We declare a pyomo model here.
model = ConcreteModel()

# Define variables
model.shifts = Var(persons, dates, within=Binary)

# each shift is assigned to 1 person.
def rule_1(model, date):
	return sum(model.shifts[person, date] for person in persons) == 1
model.each_shift_is_assigned_to_1_person = Constraint(dates, rule=rule_1)

# each person works at most one call every 3 days.
def rule_2(model, person, date):
	# list of 3 days ahead of date.
	date_list = pd.date_range(date, date + pd.DateOffset(days=2))
	if all([date in dates for date in date_list]):
		# This is not edge case.
		return sum(model.shifts[person, date] for date in date_list) <= 1
	else:
		# This is edge case. Skip the constraint.
		return Constraint.Skip
model.each_person_works_at_most_one_call_every_3_days = Constraint(persons, dates, rule=rule_2)

# No shift on annual leaves
def rule_3(model, person):
	return sum(model.shifts[person, pd.Timestamp(date)] for date in AL[person]) == 0
model.no_shift_on_annual_leave = Constraint(persons, rule=rule_3)

# Objective

# We calculate the standard deviation of everyone's shifts.
shift_counts = [sum(model.shifts[person, date] for date in dates) for person in persons]

# Mean without auxiliary variables. (Bad approach)
mean_num_shifts = sum(shift_counts) / len(shift_counts)

# Mean with auxiliary variables. (Good approach)
# mean_num_shifts = model.mean_num_shifts = Var(within=NonNegativeReals)
# model.auxiliary_constraint = Constraint(expr=mean_num_shifts == sum(shift_counts) / len(shift_counts))

standard_deviation_of_shift_counts = sum((i-mean_num_shifts)**2 for i in shift_counts) / (len(shift_counts) - 1)

model.objective = Objective(expr=standard_deviation_of_shift_counts)


# Solve

known_best_objective_value = None

def callback_function(_model):
	global known_best_objective_value, model
	if known_best_objective_value == None or value(_model.objective) < known_best_objective_value:
		
		if known_best_objective_value == None:
			print("At least one solution has been found. You may terminate anytime with Ctrl + C, or wait until the solver generates an even better solution")
		else:
			print("Better solution found.") 

		known_best_objective_value = value(_model.objective)
		for key in model.shifts:
			model.shifts[key] = _model.shifts[key]
	


try:
	SolverFactory('mindtpy').solve(model, tee=True, call_after_main_solve=callback_function)
except KeyboardInterrupt:
	pass


# Output to excel

def set_up_timetable(model, writer):
	dates = [i.strftime('%Y-%m-%d') for i in dates]

	persons_assigned = []
	for date in dates:
		persons_assigned_on_date = [person for person in persons if model.shifts[person, date].value == 1]
		persons_assigned.append(', '.join(persons_assigned_on_date))

	columns = ["Dates", "Weekday", "Call"]
	df = pd.DataFrame(list(zip(dates, persons_assigned)))
	print(df)
	df.to_excel(writer, sheet_name='Timetable', index=False)

writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')



