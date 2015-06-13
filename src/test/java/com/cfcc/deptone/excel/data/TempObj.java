package com.cfcc.deptone.excel.data;

import java.math.BigDecimal;

public class TempObj {
	private String sbtcode;
	private String name;
	private int navelamt;
	private Integer navelyearamt;
	private BigDecimal localamt;
	private Double localyearamt;
	private Double SUM_amt;
	private Double SUM_yearamt;
	private Double endamt;
	private String title;
	private String title1;
	private String title2;
	private String title3;
	private double rate=0.1;

	public Double getEndamt() {
		return endamt;
	}

	public void setEndamt(Double endamt) {
		this.endamt = endamt;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getTitle1() {
		return title1;
	}

	public void setTitle1(String title1) {
		this.title1 = title1;
	}

	public String getTitle2() {
		return title2;
	}

	public void setTitle2(String title2) {
		this.title2 = title2;
	}

	public String getTitle3() {
		return title3;
	}

	public void setTitle3(String title3) {
		this.title3 = title3;
	}

	public String getSbtcode() {
		return sbtcode;
	}

	public void setSbtcode(String sbtcode) {
		this.sbtcode = sbtcode;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getNavelamt() {
		return navelamt;
	}

	public void setNavelamt(int navelamt) {
		this.navelamt = navelamt;
	}

	public Integer getNavelyearamt() {
		return navelyearamt;
	}

	public void setNavelyearamt(Integer navelyearamt) {
		this.navelyearamt = navelyearamt;
	}

	public BigDecimal getLocalamt() {
		return localamt;
	}

	public void setLocalamt(BigDecimal localamt) {
		this.localamt = localamt;
	}

	public Double getLocalyearamt() {
		return localyearamt;
	}

	public void setLocalyearamt(Double localyearamt) {
		this.localyearamt = localyearamt;
	}

	public Double getSUM_amt() {
		return SUM_amt;
	}

	public void setSUM_amt(Double sUM_amt) {
		SUM_amt = sUM_amt;
	}

	public Double getSUM_yearamt() {
		return SUM_yearamt;
	}

	public void setSUM_yearamt(Double sUM_yearamt) {
		SUM_yearamt = sUM_yearamt;
	}

	public void setRate(double d) {
		rate = d;
	}
	public double getRate() {
		return rate;
	}
}
