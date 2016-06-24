package com.kms.convert.xml.xlsx.model;

import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.annotation.XmlAttribute;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlElementWrapper;
import javax.xml.bind.annotation.XmlRootElement;

@XmlRootElement(name = "testcase")
public class TestCase {

    private Long internalId;
    private String name;
    private Long nodeOrder;
    private Long externalId;
    private Long version;
    private String summary;
    private String preconditions;
    private Long executionType;
    private Long importance;
    private List<Step> steps = new ArrayList<Step>();

    @XmlAttribute(name = "internalid")
    public Long getInternalId() {
        return internalId;
    }

    public void setInternalId(Long internalId) {
        this.internalId = internalId;
    }

    @XmlAttribute()
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @XmlElement(name = "node_order")
    public Long getNodeOrder() {
        return nodeOrder;
    }

    public void setNodeOrder(Long nodeOrder) {
        this.nodeOrder = nodeOrder;
    }

    @XmlElement(name = "externalid")
    public Long getExternalId() {
        return externalId;
    }

    public void setExternalId(Long externalId) {
        this.externalId = externalId;
    }

    @XmlElement
    public Long getVersion() {
        return version;
    }

    public void setVersion(Long version) {
        this.version = version;
    }

    @XmlElement
    public String getSummary() {
        return summary;
    }

    public void setSummary(String summary) {
        this.summary = summary;
    }

    @XmlElement
    public String getPreconditions() {
        return preconditions;
    }

    public void setPreconditions(String preconditions) {
        this.preconditions = preconditions;
    }

    @XmlElement(name = "execution_type")
    public Long getExecutionType() {
        return executionType;
    }

    public void setExecutionType(Long executionType) {
        this.executionType = executionType;
    }

    @XmlElement
    public Long getImportance() {
        return importance;
    }

    public void setImportance(Long importance) {
        this.importance = importance;
    }

    @XmlElementWrapper(name = "steps")
    @XmlElement(name = "step")
    public List<Step> getSteps() {
        return steps;
    }

    public void setSteps(List<Step> steps) {
        this.steps = steps;
    }

}
