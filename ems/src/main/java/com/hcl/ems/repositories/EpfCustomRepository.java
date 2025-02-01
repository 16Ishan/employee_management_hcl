package com.hcl.ems.repositories;

import jakarta.persistence.EntityManager;
import jakarta.persistence.PersistenceContext;
import jakarta.transaction.Transactional;
import org.springframework.stereotype.Repository;

import java.util.Map;

@Repository
public class EpfCustomRepository {
    @PersistenceContext
    private EntityManager entityManager;
    @Transactional
    public void updateFinalDue(String staffNo, Map<String, Integer> finalDueMap) {
        StringBuilder updateQuery = new StringBuilder("update hcl.hcl.epf set final_due = CASE ");
        finalDueMap.forEach((key, value) -> {
            updateQuery.append(" WHEN month_year='").append(key)
                    .append("' THEN ").append(value)
                    .append(" ");
        });
        updateQuery.append(" END WHERE staff_no='").append(staffNo).append("'");
        entityManager.createNativeQuery(updateQuery.toString()).executeUpdate();
    }
}
