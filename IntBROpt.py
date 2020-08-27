from pulp import LpMinimize, LpProblem, LpStatus, lpSum, LpVariable
from Int_BR import IntBR


class IntBROptimizer:
    """
    Memo :
    20.08.27 Integrated BR Optimization problem

    x = Integrated BR to be calculated. (variable)
    a = 100% amount to which Integrated BR is to be applied. (constance)
    b = Reimbursement amount to which Special BR is applied. (constance)
    ax + b = Approximation of reimbursement amount using x.
    A = Actual reimbursement amount.
    z = difference between Approximation and Actual reimbursement amount.

    problem : minimize z = ax + (b - A)
    constraint(s) :
    x >= 0
    z >= 0
    """

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


class DataConstants:
    def __init__(self, plant, start_m, end_m, exc='all'):
        data = IntBR(plant, start_m, end_m)
        df = data.df
        df = df[df['WFR'] != 'F']
        if exc == 'all':
            pass
        elif exc == 'reg':
            df = df[df['일반/교류'] == 'Regular']
        elif exc == 'exc':
            df = df[df['일반/교류'] == 'Exchange']
        else :
            raise ValueError ('possible exe arguments : all, reg, exc')
        self.a = df[(df['분담율검증'] == 'Integrated BR') & (df['보상합계_기준'] != 0)]['변제대상 합계'].sum()
        self.b = df[df['분담율검증'] == 'Special BR']['보상합계_기준'].sum()
        self.A = df['변제합계_기준'].sum()


if __name__ == "__main__":
    plant = input('고객사 입력 : ')
    start_m = input('시작월 : ')
    end_m = input('종료월 : ')
    exc= input('교류클레임 선택 (1 : 전체, 2 : 일반, 3 : 교류, 기본값 : 전체)')
    if exc == '2':
        exc = 'reg'
    elif exc == '3':
        exc = 'exc'
    else:
        exc = 'all'

    Constants = DataConstants(plant, start_m, end_m, exc=exc)
    print(Constants.a, Constants.b, Constants.A)

    opt = IntBROptimizer(Constants.a, Constants.b, Constants.A)
    opt.optimize()
