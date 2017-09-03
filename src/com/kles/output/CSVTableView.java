/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.kles.output;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;
import javafx.scene.control.TableView;

/**
 *
 * @author jchau
 */
public class CSVTableView extends AbstractOutput {

    private File file;
    private TableView table;
    private String[] header;
    private List<String[]> data;

    @Override
    public void read() throws FileNotFoundException, IOException {
        Function<String, String[]> mapToData = (line) -> {
            return line.split(";");
        };
        if (file != null) {
            if (file.exists()) {
                try {
                    InputStream ins = new FileInputStream(file);
                    BufferedReader br = new BufferedReader(new InputStreamReader(ins));
                    header = br.lines().findFirst().map(mapToData).get();
                    data = br.lines().skip(1).map(mapToData).collect(Collectors.toList());
                    br.close();
                } catch (IOException e) {
                }
            }
        }
    }

    @Override
    public void write() throws FileNotFoundException, IOException {
        if (file != null && data != null) {
            Path path = Paths.get(file.getAbsolutePath());
            if (header != null) {
                try (BufferedWriter writer = Files.newBufferedWriter(path)) {
                    writer.write(String.join(";", header));
                    data.forEach(new Consumer<String[]>() {
                        @Override
                        public void accept(String[] t) {
                            try {
                                writer.write("\n" + String.join(";", t));
                            } catch (IOException ex) {
                                Logger.getLogger(CSVTableView.class.getName()).log(Level.SEVERE, null, ex);
                            }
                        }
                    });
                }
            }
        }
    }

    @Override
    public void process() throws FileNotFoundException, IOException {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public boolean save() {
        return true;
    }

    public File getFile() {
        return file;
    }

    public void setFile(File f) {
        this.file = f;
    }

    public TableView getTableView() {
        return table;
    }

    public void setTableView(TableView t) {
        this.table = t;
    }

    public String[] getHeader() {
        return header;
    }

    public void setHeader(String[] header) {
        this.header = header;
    }

    public List<String[]> getData() {
        return data;
    }

    public void setData(List<String[]> data) {
        this.data = data;
    }

}
