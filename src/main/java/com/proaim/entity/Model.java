package com.proaim.entity;

import com.sargeraswang.util.ExcelUtil.ExcelCell;
import lombok.Data;

import java.util.Date;

@Data
public class Model {
    @ExcelCell(index = 0)
    private String a;
    @ExcelCell(index = 1)
    private String b;
    @ExcelCell(index = 2)
    private String c;
    @ExcelCell(index = 3)
    private Date d;

    public Model(String a, String b, String c, Date d) {
        this.a = a;
        this.b = b;
        this.c = c;
        this.d = d;
    }
}
