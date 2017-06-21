/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.kles.output;

import java.io.FileNotFoundException;
import java.io.IOException;

/**
 *
 * @author Jeremy.CHAUT
 */
public interface IOutput {

    public void read() throws FileNotFoundException, IOException;

    public void write() throws FileNotFoundException, IOException;

    public void process() throws FileNotFoundException, IOException;

    public boolean save();
}
