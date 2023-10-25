package com.exa.beans;

import lombok.Data;

import javax.persistence.Entity;

@Data
@Entity
public class exportBean {

    private String id;

    private String name;

    private String text;
}
