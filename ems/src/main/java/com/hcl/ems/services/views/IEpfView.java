package com.hcl.ems.services.views;

public interface IEpfView
{
    String getMemberId();
    String getMonthYear();
    double getTotalSalary();
    Long getEpsOnTotalWage();
    Long getAdditionalContribution();
    Long getEpsDeposited();
    Long getEpsDue();
    double getRate();
}
