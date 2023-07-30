package com.hcl.ems.repositories;

import com.hcl.ems.model.Epf;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface EpfRepository extends JpaRepository<Epf, String>
{
    @Query("from Epf where memberId in :memberIdList and monthYear=:monthYear")
    List<Epf> findByMemberIdAndMonthYear(List<String> memberIdList, String monthYear);
}