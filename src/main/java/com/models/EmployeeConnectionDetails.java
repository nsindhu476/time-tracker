package com.models;


import javax.validation.constraints.NotBlank;

public class EmployeeConnectionDetails {
    @NotBlank
    private String lid;
    private String email;
    private String contype;
    
    public EmployeeConnectionDetails(@NotBlank String lid, String email, String contype) {
		super();
		this.lid = lid;
		this.email = email;
		this.contype = contype;
	}
	public String getLid() {
		return lid;
	}
	public void setLid(String lid) {
		this.lid = lid;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public String getContype() {
		return contype;
	}
	public void setContype(String contype) {
		this.contype = contype;
	}
	@Override
	public String toString() {
		return "EmployeeConnectionDetails [lid=" + lid + ", email=" + email + ", contype=" + contype + "]";
	}
	public EmployeeConnectionDetails()
	{
		
	}

   
}

