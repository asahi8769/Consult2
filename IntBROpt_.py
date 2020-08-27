from pulp import LpMinimize, LpProblem, LpStatus, lpSum, LpVariable


class IntBROptimizer:
    def __init__(self, a_int100, b_indapp, A_reimb):
        self.a = a_int100
        self.b = b_indapp
        self.A = A_reimb
        self.check_type()
        self.x = LpVariable(name="x", lowBound=0, upBound=1, cat='Continuous')
        self.model = LpProblem(name='Int_BR_Optimization_problem', sense=LpMinimize)
        self.z = self.a * self.x + self.b - self.A
        self.const_1 = self.z >= 0
        self.model += self.z
        self.model += (self.const_1, 'constrain_1')
        self.target = None
        print(self.model)

    def check_type(self):
        for variable in (self.a, self.b, self.A):
            if type(variable) == str:
                raise AssertionError('a_int100, b_indapp, A_reimb should be numeric.')

    def optimize(self):
        status = self.model.solve()
        self.target = round(self.model.variables()[0].value(), 4)
        print(f"Status: {status}")
        print(f"Objective: {round(self.model.objective.value(),4)}")
        print(f"Target Value: {self.target}")


if __name__ == "__main__":
    opt = IntBROptimizer(354, 100, 205)
    opt.optimize()
    print(opt.target)
