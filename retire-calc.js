// Retirement Planner calculations.
//
// Ported line for line from Jake Ippolito's Retirement Planning Calculator
// (github.com/jakeipp/retirement-planner). Only names and formatting were made
// readable; every formula, limit and default is unchanged. If the original
// changes, update this file to match.

// Social Security: sliding scale on final-year income, adjusted for
// retirement age, x1.5 for a married couple, capped at the maximum.
const SS_INCOME_FOR_MAX = 176100;
const SS_MAX_BENEFIT = 48216;
const SS_FULL_RETIREMENT_AGE = 67;

function socialSecurityBenefit(income, isMarried, retirementAge) {
  let benefit = Math.min(income / SS_INCOME_FOR_MAX, 1) * SS_MAX_BENEFIT;
  if (retirementAge < SS_FULL_RETIREMENT_AGE) {
    const yearsEarly = SS_FULL_RETIREMENT_AGE - retirementAge;
    benefit *= 1 - yearsEarly * 0.067;
  } else if (retirementAge > SS_FULL_RETIREMENT_AGE) {
    const yearsLate = Math.min(retirementAge - SS_FULL_RETIREMENT_AGE, 3);
    benefit *= 1 + yearsLate * 0.08;
  }
  if (isMarried) benefit *= 1.5;
  const cap = isMarried ? SS_MAX_BENEFIT * 1.5 : SS_MAX_BENEFIT;
  return Math.min(benefit, cap);
}

// Year-by-year projection. No contribution is made in the year of retirement.
function projectRetirement(e) {
  const {
    currentAge, retirementAge, annualIncome, percentToSave, currentSavings,
    expectedIncomeIncrease, preRetirementIncomeDesired, yearsOfRetirementIncome,
    returnBeforeRetirement, returnDuringRetirement, expectedInflation,
    isMarried, includeSocialSecurity,
  } = e;
  const yearlyData = [];
  let balance = currentSavings;
  let income = annualIncome;
  let lastYearIncome = annualIncome;
  let savingsRunOutAge = retirementAge + yearsOfRetirementIncome;
  const endAge = retirementAge + yearsOfRetirementIncome;

  for (let age = currentAge + 1; age < retirementAge; age++) {
    const beginningBalance = balance;
    const investmentGrowth = beginningBalance * (returnBeforeRetirement / 100);
    income = income * (1 + expectedIncomeIncrease / 100);
    lastYearIncome = income;
    const contributions = income * (percentToSave / 100);
    balance = beginningBalance + investmentGrowth + contributions;
    yearlyData.push({ age, beginningBalance, investmentGrowth, contributions, annualIncomeDesired: 0, withdrawals: 0, endingBalance: balance });
  }
  if (retirementAge > currentAge + 1) {
    income = income * (1 + expectedIncomeIncrease / 100);
    lastYearIncome = income;
  }
  const retirementBalance = balance;
  const annualRetirementExpenses = lastYearIncome * (preRetirementIncomeDesired / 100);
  const annualSocialSecurity = includeSocialSecurity ? socialSecurityBenefit(lastYearIncome, isMarried, retirementAge) : 0;

  let incomeDesired = annualRetirementExpenses;
  for (let age = retirementAge; age < endAge; age++) {
    const beginningBalance = balance;
    const investmentGrowth = beginningBalance * (returnDuringRetirement / 100);
    if (age > retirementAge) incomeDesired = incomeDesired * (1 + expectedInflation / 100);
    const needed = Math.max(0, incomeDesired - annualSocialSecurity);
    const withdrawals = Math.min(needed, beginningBalance + investmentGrowth);
    balance = Math.max(0, beginningBalance + investmentGrowth - withdrawals);
    if (balance === 0 && savingsRunOutAge === endAge) savingsRunOutAge = age;
    yearlyData.push({ age, beginningBalance, investmentGrowth, contributions: 0, annualIncomeDesired: incomeDesired, withdrawals, endingBalance: balance });
  }
  return { retirementBalance, lastYearIncome, annualRetirementExpenses, annualSocialSecurity, savingsRunOutAge, endingBalance: balance, yearlyData };
}

const lastsToGoal = (inputs) => projectRetirement(inputs).savingsRunOutAge >= inputs.retirementAge + inputs.yearsOfRetirementIncome;

// Savings rate needed (binary search 0-100%, rounded up to 0.01).
function recommendedSavingsRate(e) {
  let lo = 0, hi = 100;
  for (let i = 0; i < 50; i++) {
    const mid = (lo + hi) / 2;
    if (lastsToGoal({ ...e, percentToSave: mid })) hi = mid; else lo = mid;
  }
  return Math.ceil(hi * 100) / 100;
}

// Pre-retirement return needed (binary search 0-20%, rounded up to 0.01).
function recommendedReturnRate(e) {
  let lo = 0, hi = 20;
  for (let i = 0; i < 50; i++) {
    const mid = (lo + hi) / 2;
    if (lastsToGoal({ ...e, returnBeforeRetirement: mid })) hi = mid; else lo = mid;
  }
  return Math.ceil(hi * 100) / 100;
}

// Retirement income the savings can support (binary search 40%-current, rounded down).
function recommendedIncomePercent(e) {
  let lo = 40, hi = e.preRetirementIncomeDesired;
  for (let i = 0; i < 50; i++) {
    const mid = (lo + hi) / 2;
    if (lastsToGoal({ ...e, preRetirementIncomeDesired: mid })) lo = mid; else hi = mid;
  }
  return Math.floor(lo);
}

// First retirement age (up to 90) at which the savings last.
function recommendedRetirementAge(e) {
  for (let age = e.retirementAge; age <= 90; age++) {
    if (lastsToGoal({ ...e, retirementAge: age })) return age;
  }
  return 90;
}

function calculateRetirement(e) {
  return {
    ...projectRetirement(e),
    recommendedSavingsRate: recommendedSavingsRate(e),
    recommendedReturnRate: recommendedReturnRate(e),
    recommendedIncomePercent: recommendedIncomePercent(e),
    recommendedRetirementAge: recommendedRetirementAge(e),
  };
}

const RETIREMENT_DEFAULTS = {
  currentAge: 25, retirementAge: 65, annualIncome: 100000, percentToSave: 5, currentSavings: 0,
  brokerageAccount: 0, traditionalIraAccount: 0, rothIraAccount: 0, account401k: 0, otherAccounts: 0,
  expectedIncomeIncrease: 0, preRetirementIncomeDesired: 40, yearsOfRetirementIncome: 25,
  returnBeforeRetirement: 10, returnDuringRetirement: 5, expectedInflation: 0,
  isMarried: false, includeSocialSecurity: false,
  userInfo: { name: "", email: "", phoneNumber: "" },
};
