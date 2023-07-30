package com.hcl.ems.validators;

import com.hcl.ems.dto.CellDto;
import com.hcl.ems.model.Epf;
import org.springframework.stereotype.Component;

import java.util.Objects;
import java.util.concurrent.atomic.AtomicBoolean;

@Component
public class DataValidator
{
    public void verifySnoSlipMemberNameCatStaff(Epf epf, Epf epfMismatch, CellDto cellDto,
                                                 AtomicBoolean mismatch)
    {
        if(!Objects.equals(epf.getSNo(), cellDto.getSNo()))
        {
            mismatch.set(true);
            epfMismatch.setSNo(cellDto.getSNo());
        }
        if(!Objects.equals(epf.getSlipNo(), cellDto.getSlipNo()))
        {
            mismatch.set(true);
            epfMismatch.setSlipNo(cellDto.getSlipNo());
        }
        if(!epf.getMemberId().equals(cellDto.getMemberId()))
        {
            mismatch.set(true);
            epfMismatch.setMemberId(cellDto.getMemberId());
        }
        if(!epf.getName().equals(cellDto.getName()))
        {
            mismatch.set(true);
            epfMismatch.setName(cellDto.getName());
        }
        if(!epf.getCategory().equals(cellDto.getCategory()))
        {
            mismatch.set(true);
            epfMismatch.setCategory(cellDto.getCategory());
        }
        if(!Objects.equals(epf.getStaffNo(), cellDto.getStaffNo()))
        {
            mismatch.set(true);
            epfMismatch.setStaffNo(cellDto.getStaffNo());
        }
    }

    public void verifyBasicPpDaBasicArrPpArrDaArr(Epf epf, Epf epfMismatch, CellDto cellDto,
                                                   AtomicBoolean mismatch)
    {
        if(!Objects.equals(epf.getBasic(), cellDto.getBasic()))
        {
            mismatch.set(true);
            epfMismatch.setBasic(cellDto.getBasic());
        }
        if(!Objects.equals(epf.getPp(), cellDto.getPp()))
        {
            mismatch.set(true);
            epfMismatch.setPp(cellDto.getPp());
        }
        if(!Objects.equals(epf.getDaVda(), cellDto.getDaVda()))
        {
            mismatch.set(true);
            epfMismatch.setDaVda(cellDto.getDaVda());
        }
        if(!Objects.equals(epf.getBasicArrears(), cellDto.getBasicArrears()))
        {
            mismatch.set(true);
            epfMismatch.setBasicArrears(cellDto.getBasicArrears());
        }
        if(!Objects.equals(epf.getPpArrears(), cellDto.getPpArrears()))
        {
            mismatch.set(true);
            epfMismatch.setPpArrears(cellDto.getPpArrears());
        }
        if(!Objects.equals(epf.getDaArrears(), cellDto.getDaArrears()))
        {
            mismatch.set(true);
            epfMismatch.setDaArrears(cellDto.getDaArrears());
        }
    }

    public void verifyCpfArrEpsArrTotalSalCpfCpfArrDedTotalCpf(Epf epf, Epf epfMismatch,
                                                                CellDto cellDto, AtomicBoolean mismatch)
    {
        if(!Objects.equals(epf.getCpfArrears(), cellDto.getCpfArrears()))
        {
            mismatch.set(true);
            epfMismatch.setCpfArrears(cellDto.getCpfArrears());
        }
        if(!Objects.equals(epf.getEpsArrears(), cellDto.getEpsArrears()))
        {
            mismatch.set(true);
            epfMismatch.setEpsArrears(cellDto.getEpsArrears());
        }
        if(!Objects.equals(epf.getTotalSalary(), cellDto.getTotalSalary()))
        {
            mismatch.set(true);
            epfMismatch.setTotalSalary(cellDto.getTotalSalary());
        }
        if(!Objects.equals(epf.getCpf(), cellDto.getCpf()))
        {
            mismatch.set(true);
            epfMismatch.setCpf(cellDto.getCpf());
        }
        if(!Objects.equals(epf.getCpfArrearsDed(), cellDto.getCpfArrearsDed()))
        {
            mismatch.set(true);
            epfMismatch.setCpfArrearsDed(cellDto.getCpfArrearsDed());
        }
        if(!Objects.equals(epf.getTotalCpf(), cellDto.getTotalCpf()))
        {
            mismatch.set(true);
            epfMismatch.setTotalCpf(cellDto.getTotalCpf());
        }
    }

    public void verifyEpsEpsArrDedTotalEpsRemarkMonth(Epf epf, Epf epfMismatch, CellDto cellDto,
                                                       AtomicBoolean mismatch, String month)
    {
        if(!Objects.equals(epf.getEps(), cellDto.getEps()))
        {
            mismatch.set(true);
            epfMismatch.setEps(cellDto.getEps());
        }
        if(!Objects.equals(epf.getEpsArrearsDed(), cellDto.getEpsArrearsDed()))
        {
            mismatch.set(true);
            epfMismatch.setEpsArrearsDed(cellDto.getEpsArrearsDed());
        }
        if(!Objects.equals(epf.getTotalEps(), cellDto.getTotalEps()))
        {
            mismatch.set(true);
            epfMismatch.setTotalEps(cellDto.getTotalEps());
        }
        if(!epf.getRemark().equals(cellDto.getRemark()))
        {
            mismatch.set(true);
            epfMismatch.setRemark(cellDto.getRemark());
        }
        if(!epf.getMonthYear().equals(month))
        {
            mismatch.set(true);
            epfMismatch.setMonthYear(month);
        }
    }
}
