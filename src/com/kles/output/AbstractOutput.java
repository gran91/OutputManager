/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.kles.output;

import java.io.IOException;

/**
 *
 * @author Jeremy.CHAUT
 */
public abstract class AbstractOutput implements IOutput, Runnable {

    protected String action;
    protected String filepath;
    public static final String READ = "READ";
    public static final String WRITE = "WRITE";
    public static final String COPYFILE = "COPYFILE";
    public static final String DELETEFILE = "DELETEFILE";

    @Override
    public void run() {
        switch (action) {
            case WRITE:
                try {
                    write();
                    save();
                } catch (IOException ex) {
                }
                break;
            case READ:
                try {
                    write();
                    save();
                } catch (IOException ex) {
                }
                break;
        }
    }

    public String getAction() {
        return action;
    }

    public void setAction(String action) {
        this.action = action;
    }

    public String getFilePath() {
        return filepath;
    }

    public void setFilePath(String filepath) {
        this.filepath = filepath;
    }
}
