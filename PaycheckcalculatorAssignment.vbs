Dim annualSalary, taxes, retContribution, healthPlan

Dim name, mothlySalary , deductions, netpay

name = InputBox ("Enter name of employee")

annualSalary = 60000
taxes = (annualSalary/100) * 20
retContribution = (annualSalary/100) * 5
healthPlan = 200
deductions = taxes + retContribution + healthPlan
netpay = (annualSalary - deductions) / 12

MsgBox name & " receives $" & netpay & " as monthly salary after all deductions! ", 0

