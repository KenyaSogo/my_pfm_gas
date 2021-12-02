/** @OnlyCurrentDoc */

function executeRetroactive1203(){
  aggregateByMonth(2); // 09/25~
  aggregateByMonth(1); // 10/25~
  aggregateByMonth(0); // 11/25~
  calcDailySummary(2); // 09/25~
  calcDailySummary(1); // 10/25~
  calcDailySummary(0); // 11/25~
  calcMonthlySummary(3); // 09/01~
  calcMonthlySummary(2); // 10/01~
  calcMonthlySummary(1); // 11/01~
  calcMonthlySummary(0); // 12/01~
  executeUpdateGraphSourceData();
}
