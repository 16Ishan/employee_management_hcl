package com.hcl.ems.repositories;

import com.hcl.ems.model.Epf;
import com.hcl.ems.services.views.IEpfView;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface EpfRepository extends JpaRepository<Epf, String>
{
    @Query("from Epf where memberId in :memberIdList and monthYear=:monthYear")
    List<Epf> findByMemberIdAndMonthYear(List<String> memberIdList, String monthYear);

    @Query("select memberId as memberId, monthYear as monthYear, totalSalary as totalSalary, " +
            "epsOnTotalWage as epsOnTotalWage, additionalContribution as additionalContribution, " +
            "epsDeposited as epsDeposited, epsDue as epsDue, rate as rate " +
            "from Epf where staffNo=:staffNo " +
            "ORDER BY CAST(RIGHT(monthYear, 4) AS INTEGER), CAST(LEFT(monthYear, 2) AS INTEGER)")
    List<IEpfView> findByStaffNo(String staffNo);

    @Query("from Epf where staffNo=:staffNo ORDER BY CAST(RIGHT(monthYear, 4) AS INTEGER), CAST(LEFT(monthYear, 2) AS INTEGER)")
    List<Epf> findAllByStaffNo(String staffNo);
}