/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.kles.output;

import java.io.FileNotFoundException;
import java.io.IOException;
import javafx.concurrent.Task;

/**
 *
 * @author Jeremy.CHAUT
 */
public abstract class AbstractOutput extends Task<Void> implements IOutput, Runnable {

    protected String action;
    protected String filepath;
    public static final String READ = "READ";
    public static final String WRITE = "WRITE";
    public static final String COPYFILE = "COPYFILE";
    public static final String DELETEFILE = "DELETEFILE";
    public static final String PROCESS = "PROCESS";
    public static final String SAVE = "SAVE";

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

    @Override
    protected Void call() throws Exception {
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
                    read();
                } catch (IOException ex) {
                }
                break;
            case PROCESS:
                try {
                    process();
                } catch (IOException ex) {
                }
                break;
            case SAVE:
                save();
                break;
        }
        return null;
    }

    @Override
    public void read() throws FileNotFoundException, IOException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void write() throws FileNotFoundException, IOException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void process() throws FileNotFoundException, IOException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean save() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
}
