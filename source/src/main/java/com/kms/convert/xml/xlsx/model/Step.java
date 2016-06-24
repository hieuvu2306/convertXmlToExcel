package com.kms.convert.xml.xlsx.model;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;

@XmlRootElement(name = "step")
public class Step {
    private Long stepNumber;
    private String actions;
    private String expectedResults;
    private Long executionType;

    public Long getStepNumber() {
        return stepNumber;
    }

    @XmlElement(name = "step_number")
    public void setStepNumber(Long stepNumber) {
        this.stepNumber = stepNumber;
    }

    public String getActions() {
        return actions;
    }

    @XmlElement
    public void setActions(String actions) {
        this.actions = actions;
    }

    @XmlElement(name = "expectedresults")
    public String getExpectedResults() {
        return expectedResults;
    }

    public void setExpectedResults(String expectedResults) {
        this.expectedResults = expectedResults;
    }

    @XmlElement(name = "execution_type")
    public Long getExecutionType() {
        return executionType;
    }

    public void setExecutionType(Long executionType) {
        this.executionType = executionType;
    }

}
