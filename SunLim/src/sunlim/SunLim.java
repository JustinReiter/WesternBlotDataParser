/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template allF, choose Tools | Templates
 * and open the template in the editor.
 */
package sunlim;

import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;

import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author Matthew
 */
public class SunLim {

    public static BufferedReader read = null;
    public static PrintWriter write = null;

    public static JFrame frame = new JFrame("Sun-Lim's Program");
    public static JPanel mainPanel = new JPanel(new GridBagLayout());
    public static JPanel leftBut = new JPanel(new GridBagLayout());
    public static JPanel rightBut = new JPanel(new GridBagLayout());
    public static JPanel rightTex = new JPanel(new GridBagLayout());

    public static GridBagConstraints gbConstraints = new GridBagConstraints();

    public static JFileChooser fc = new JFileChooser();
    public static JButton allFiles = new JButton("Set Main File w/ Everything");
    public static JButton targ1 = new JButton("Set First Target");
    public static JButton targ2 = new JButton("Set Second Target");
    public static JButton targ3 = new JButton("Set Third Target");
    public static JButton targ4 = new JButton("Set Fourth Target");
    public static JButton ctrl = new JButton("Set Control");
    public static JButton out = new JButton("Set Output");
    public static JButton run = new JButton("RUN");

    public static JLabel incLabel = new JLabel("Set the increment for results:");
    public static JTextField incText = new JTextField("3");
    public static JLabel ctrlLabel = new JLabel("What is the name of the control antibody?");
    public static JTextField ctrlText = new JTextField("bActin");

    private static File allF;
    private static File fileT1 = null;
    private static File fileT2 = null;
    private static File fileT3 = null;
    private static File fileT4 = null;
    private static File fileC = null;

    private static File fileAF = new File("");
    private static File fileTF1 = new File("");
    private static File fileTF2 = new File("");
    private static File fileTF3 = new File("");
    private static File fileTF4 = new File("");
    private static File fileCF = new File("");

    private static File save = new File("all.txt");
    private static File saveT = new File("");

    public static void componentSetter(JComponent component, int grdX, int grdY, int gWidth, int gHeight, double wetX, double wetY) {
        gbConstraints.gridx = grdX;
        gbConstraints.gridy = grdY;
        gbConstraints.gridwidth = gWidth;
        gbConstraints.gridheight = gHeight;
        gbConstraints.weightx = wetX;
        gbConstraints.weighty = wetY;
        gbConstraints.insets = new Insets(0, 10, 0, 0);
        gbConstraints.fill = GridBagConstraints.BOTH;
        mainPanel.add(component, gbConstraints);
    }

    public static void main(String[] args) throws IOException {
        fileTF1.deleteOnExit();
        fileTF2.deleteOnExit();
        fileTF3.deleteOnExit();
        fileTF4.deleteOnExit();
        fileCF.deleteOnExit();
        fileAF.deleteOnExit();
        saveT.deleteOnExit();

        targ1.addActionListener(new ButtonAction());
        targ1.setActionCommand("targ1");
        targ2.addActionListener(new ButtonAction());
        targ2.setActionCommand("targ2");
        targ3.addActionListener(new ButtonAction());
        targ3.setActionCommand("targ3");
        targ4.addActionListener(new ButtonAction());
        targ4.setActionCommand("targ4");
        ctrl.addActionListener(new ButtonAction());
        ctrl.setActionCommand("ctrl");
        allFiles.addActionListener(new ButtonAction());
        allFiles.setActionCommand("allFiles");
        out.addActionListener(new ButtonAction());
        out.setActionCommand("output");
        run.addActionListener(new ButtonAction());
        run.setActionCommand("run");

        componentSetter(targ1, 0, 0, 1, 1, 0.33, 0.2);
        componentSetter(targ2, 0, 1, 1, 1, 0.33, 0.2);
        componentSetter(targ3, 0, 2, 1, 1, 0.33, 0.2);
        componentSetter(targ4, 0, 3, 1, 1, 0.33, 0.2);
        componentSetter(ctrl, 1, 0, 1, 1, 0.33, 0.2);
        componentSetter(allFiles, 1, 1, 1, 1, 0.33, 0.2);
        componentSetter(out, 1, 2, 1, 1, 0.33, 0.2);
        componentSetter(run, 1, 3, 1, 1, 0.33, 0.2);
        componentSetter(incLabel, 2, 0, 1, 1, 0.33, 0.2);
        componentSetter(incText, 2, 1, 1, 1, 0.33, 0.2);
        componentSetter(ctrlLabel, 2, 2, 1, 1, 0.33, 0.2);
        componentSetter(ctrlText, 2, 3, 1, 1, 0.33, 0.2);

        frame.setPreferredSize(new Dimension(800, 300));
        frame.setVisible(true);
        frame.setResizable(false);
        frame.setContentPane(mainPanel);
        frame.pack();
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        try {
            UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
            SwingUtilities.updateComponentTreeUI(frame);
            SwingUtilities.updateComponentTreeUI(mainPanel);
            SwingUtilities.updateComponentTreeUI(fc);
        } catch (Exception e) {
            System.out.println(("Stupid Look and Feel"));
        }
    }

    public static class ButtonAction implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent e) {
            String act = e.getActionCommand();
            JButton but = (JButton) e.getSource();
            System.out.println("HI");
            try {
                if (act.equals("run")) {
                    try {
                        if (allF != null) {
                            fileAF = convertExcel(allF);
                            fileAF.deleteOnExit();
                            runAll(fileAF.getAbsolutePath());
                            convertCSV();
                        } else {
                            System.out.println("NEXT");
                            fileTF2 = new File("");
                            fileTF3 = new File("");
                            fileTF4 = new File("");
                            fileCF = new File("");
                            fileTF1 = convertExcel(fileT1);
                            fileTF1.deleteOnExit();
                            if (!(fileT2 == null)) {
                                fileTF2 = convertExcel(fileT2);
                                fileTF2.deleteOnExit();
                            }
                            if (!(fileT3 == null)) {
                                fileTF3 = convertExcel(fileT3);
                                fileTF3.deleteOnExit();
                            }
                            if (!(fileT4 == null)) {
                                fileTF4 = convertExcel(fileT4);
                                fileTF4.deleteOnExit();
                            }
                            if (!(fileC == null)) {
                                fileCF = convertExcel(fileC);
                                fileCF.deleteOnExit();
                            }
                            runInd(fileTF1.getAbsolutePath(), fileTF2.getAbsolutePath(), fileTF3.getAbsolutePath(), fileTF4.getAbsolutePath(), fileCF.getAbsolutePath());
                            convertCSV();
                        }
                    } catch (Exception err) {
                        System.out.println("HI1234");
                    } finally {
                        if (read != null) {
                            read.close();
                        }
                        if (write != null) {
                            write.close();
                        }
                    }
                    System.out.println("HI2");
                } else if (act.indexOf("ctrl") != -1) {
                    int returnValC = fc.showOpenDialog(mainPanel);
                    if (returnValC == JFileChooser.APPROVE_OPTION) {
                        fileC = fc.getSelectedFile();
                        //This is where a real application would open the allF.
                        System.out.println("Opening: " + fileC.getName() + ".\n");
                    } else {
                        System.out.println("Open command cancelled by user.\n");
                    }
                } else if (act.indexOf("output") != -1) {
                    int saveVal = fc.showSaveDialog(mainPanel);
                    if (saveVal == JFileChooser.APPROVE_OPTION) {
                        save = fc.getSelectedFile();
                        System.out.println("Success");
                    }
                } else if (act.indexOf("targ") != -1) {
                    // MESSY PLEASE FIX
                    switch (act.substring(4)) {
                        case ("1"):
                            int returnVal1 = fc.showOpenDialog(mainPanel);
                            if (returnVal1 == JFileChooser.APPROVE_OPTION) {
                                fileT1 = fc.getSelectedFile();
                                //This is where a real application would open the allF.
                                System.out.println("Opening: " + fileT1.getName() + ".\n");
                            } else {
                                System.out.println("Open command cancelled by user.\n");
                            }
                            break;
                        case ("2"):
                            int returnVal2 = fc.showOpenDialog(mainPanel);
                            if (returnVal2 == JFileChooser.APPROVE_OPTION) {
                                fileT2 = fc.getSelectedFile();
                                //This is where a real application would open the allF.
                                System.out.println("Opening: " + fileT2.getName() + ".\n");
                            } else {
                                System.out.println("Open command cancelled by user.\n");
                            }
                            break;
                        case ("3"):
                            int returnVal3 = fc.showOpenDialog(mainPanel);
                            if (returnVal3 == JFileChooser.APPROVE_OPTION) {
                                fileT3 = fc.getSelectedFile();
                                //This is where a real application would open the allF.
                                System.out.println("Opening: " + fileT3.getName() + ".\n");
                            } else {
                                System.out.println("Open command cancelled by user.\n");
                            }
                            break;
                        default:
                            int returnVal4 = fc.showOpenDialog(mainPanel);
                            if (returnVal4 == JFileChooser.APPROVE_OPTION) {
                                fileT4 = fc.getSelectedFile();
                                //This is where a real application would open the allF.
                                System.out.println("Opening: " + fileT4.getName() + ".\n");
                            } else {
                                System.out.println("Open command cancelled by user.\n");
                            }
                            break;
                    }
                } else if (act.indexOf("all") != -1) {
                    int returnVal = fc.showOpenDialog(mainPanel);

                    if (returnVal == JFileChooser.APPROVE_OPTION) {
                        allF = fc.getSelectedFile();
                        //This is where a real application would open the allF.
                        System.out.println("Opening: " + allF.getName() + ".\n");
                    } else {
                        System.out.println("Open command cancelled by user.\n");
                    }
                    System.out.println("HI3");
                }
            } catch (Exception exception) {
                System.out.println("HI1234");
                exception.printStackTrace();
            }
        }
    }

    public static void setFile() {
        int returnVal = fc.showOpenDialog(mainPanel);

        if (returnVal == JFileChooser.APPROVE_OPTION) {
            allF = fc.getSelectedFile();
            //This is where a real application would open the allF.
            System.out.println("Opening: " + allF.getName() + ".\n");
        } else {
            System.out.println("Open command cancelled by user.\n");
        }
    }

    // Method that computes the data using multiple single-table files
    public static void runInd(String targ1Str, String targ2Str, String targ3Str, String targ4Str, String ctrlStr) throws IOException {
        ArrayList<String> targ1Full = new ArrayList<>();
        ArrayList<String[]> targ1Sep = new ArrayList<>();
        ArrayList<String> targ2Full = new ArrayList<>();
        ArrayList<String[]> targ2Sep = new ArrayList<>();
        ArrayList<String> targ3Full = new ArrayList<>();
        ArrayList<String[]> targ3Sep = new ArrayList<>();
        ArrayList<String> targ4Full = new ArrayList<>();
        ArrayList<String[]> targ4Sep = new ArrayList<>();
        ArrayList<String> ctrlFull = new ArrayList<>();
        ArrayList<String[]> ctrlSep = new ArrayList<>();

        try {
            read = new BufferedReader(new FileReader(targ1Str));
            String line;
            do {
                line = read.readLine();
                if (line != null) {
                    targ1Full.add(line);
                }
            } while (line != null);

            for (int i = 0; i < targ1Full.size(); i++) {
                targ1Sep.add(targ1Full.get(i).split("\t"));
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (read != null) {
                read.close();
            }
            if (write != null) {
                write.close();
            }
        }

        if (!targ2Str.equals("") && targ2Str.substring(targ2Str.length() - 3, targ2Str.length()).equals("txt")) {
            try {
                read = new BufferedReader(new FileReader(targ2Str));
                String line;
                do {
                    line = read.readLine();
                    if (line != null) {
                        targ2Full.add(line);
                    }
                } while (line != null);

                for (int i = 0; i < targ2Full.size(); i++) {
                    targ2Sep.add(targ2Full.get(i).split("\t"));
                }

            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (read != null) {
                    read.close();
                }
                if (write != null) {
                    write.close();
                }
            }
        }

        if (!targ3Str.equals("") && targ3Str.substring(targ3Str.length() - 3, targ3Str.length()).equals("txt")) {
            try {
                read = new BufferedReader(new FileReader(targ3Str));
                String line;
                do {
                    line = read.readLine();
                    if (line != null) {
                        targ3Full.add(line);
                    }
                } while (line != null);

                for (int i = 0; i < targ3Full.size(); i++) {
                    targ3Sep.add(targ3Full.get(i).split("\t"));
                }

            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (read != null) {
                    read.close();
                }
                if (write != null) {
                    write.close();
                }
            }
        }

        if (!targ4Str.equals("") && targ4Str.substring(targ4Str.length() - 3, targ4Str.length()).equals("txt")) {
            try {
                read = new BufferedReader(new FileReader(targ4Str));
                String line;
                do {
                    line = read.readLine();
                    if (line != null) {
                        targ4Full.add(line);
                    }
                } while (line != null);

                for (int i = 0; i < targ4Full.size(); i++) {
                    targ4Sep.add(targ4Full.get(i).split("\t"));
                }

            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (read != null) {
                    read.close();
                }
                if (write != null) {
                    write.close();
                }
            }
        }

        try {
            read = new BufferedReader(new FileReader(ctrlStr));
            String line;
            do {
                line = read.readLine();
                if (line != null) {
                    ctrlFull.add(line);
                }
            } while (line != null);

            for (int i = 0; i < ctrlFull.size(); i++) {
                ctrlSep.add(ctrlFull.get(i).split("\t"));
            }
            System.out.println("FINI");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (read != null) {
                read.close();
            }
            if (write != null) {
                write.close();
            }
        }
        System.out.println(1);
        ArrayList<Double> values = new ArrayList<>();
        // Control Values
        for (int i = 2; i < ctrlSep.size(); i++) {
            values.add(Double.parseDouble(ctrlSep.get(i)[5].replace(",", "")));
        }
        System.out.println(2);
        // Targ1 *Mandatory*
        for (int i = 2; i < targ1Sep.size(); i++) {
            values.add(Double.parseDouble(targ1Sep.get(i)[5].replace(",", "")));
        }
        System.out.println(3);
        // Targ2 *Optional*
        if (targ2Sep.size() > 1) {
            for (int i = 2; i < targ2Sep.size(); i++) {
                values.add(Double.parseDouble(targ2Sep.get(i)[5].replace(",", "")));
            }
        }
        System.out.println(4);
        // Targ3 *Optional*
        if (targ3Sep.size() > 1) {
            for (int i = 2; i < targ3Sep.size(); i++) {
                values.add(Double.parseDouble(targ3Sep.get(i)[5].replace(",", "")));
            }
        }
        System.out.println(5);
        // Targ4 *Optional*
        if (targ4Sep.size() > 1) {
            for (int i = 2; i < targ4Sep.size(); i++) {
                values.add(Double.parseDouble(targ4Sep.get(i)[5].replace(",", "")));
            }
        }
        System.out.println(6);
        ArrayList<Double> calc = new ArrayList<>();

        int targInd = values.size() / (targ4Str.substring(targ4Str.length() - 3, targ4Str.length()).equals("txt") ? 5 : targ3Str.substring(targ3Str.length() - 3, targ3Str.length()).equals("txt") ? 4 : targ2Str.substring(targ2Str.length() - 3, targ2Str.length()).equals("txt") ? 3 : 2);

        for (int i = targInd; i < values.size(); i++) {
            calc.add(values.get(i) / values.get(i % targInd));
        }

        try {
            System.out.println("HI 555");
            saveT = new File(save.getAbsolutePath().replace("xlsx", "txt"));
            saveT.deleteOnExit();
            write = new PrintWriter(saveT);

            for (int i = 0; i < targ1Full.size(); i++) {
                write.println(targ1Full.get(i));
            }
            for (int i = 0; i < targInd & i < calc.size(); i++) {
                write.println((i % Double.parseDouble(incText.getText()) == 0 ? "ctrl" : "\t") + "\t\t\t\t\t" + Math.round(calc.get(i) * 100.0) / 100.0 + "\t" + Math.round((calc.get(i) / calc.get((int) (i - i % Double.parseDouble(incText.getText())))) * 10000.0) / 100.0);
            }

            for (int i = 0; i < targ2Full.size(); i++) {
                write.println(targ2Full.get(i));
            }
            for (int i = targInd; i < targInd * 2 & i < calc.size(); i++) {
                write.println((i % Double.parseDouble(incText.getText()) == 0 ? "ctrl" : "\t") + "\t\t\t\t\t" + Math.round(calc.get(i) * 100.0) / 100.0 + "\t" + Math.round((calc.get(i) / calc.get((int) (i - i % targInd % Double.parseDouble(incText.getText())))) * 10000.0) / 100.0);
            }

            for (int i = 0; i < targ3Full.size(); i++) {
                write.println(targ3Full.get(i));
            }
            for (int i = targInd * 2; i < targInd * 3 & i < calc.size(); i++) {
                write.println((i % Double.parseDouble(incText.getText()) == 0 ? "ctrl" : "\t") + "\t\t\t\t\t" + Math.round(calc.get(i) * 100.0) / 100.0 + "\t" + Math.round((calc.get(i) / calc.get((int) (i - i % targInd % Double.parseDouble(incText.getText())))) * 10000.0) / 100.0);
            }

            for (int i = 0; i < targ4Full.size(); i++) {
                write.println(targ4Full.get(i));
            }
            for (int i = targInd * 3; i < targInd * 4 & i < calc.size(); i++) {
                write.println((i % Double.parseDouble(incText.getText()) == 0 ? "ctrl" : "\t") + "\t\t\t\t\t" + Math.round(calc.get(i) * 100.0) / 100.0 + "\t" + Math.round((calc.get(i) / calc.get((int) (i - i % targInd % Double.parseDouble(incText.getText())))) * 10000.0) / 100.0);
            }

            for (int i = 0; i < ctrlFull.size(); i++) {
                write.println(ctrlFull.get(i));
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (write != null) {
                write.close();
            }
        }
    }

    // Method that computes the data using file with all tables
    public static void runAll(String fileStr) throws IOException {
        try {
            read = new BufferedReader(new FileReader(fileStr));
            String line;
            ArrayList<String> store = new ArrayList<>();
            ArrayList<String[]> sep = new ArrayList<>();
            store.add("");
            do {
                line = read.readLine();
                if (line != null) {
                    store.add(line);
                } else if (line == null) {
                    break;
                } else {
                    store.add("");
                }
            } while (line != null);

            for (int i = 0; i < store.size(); i++) {
                sep.add(store.get(i).split("\t"));
            }

            int ctrlInd = 0;
            int expInd = 0;
            ArrayList<Double> comp = new ArrayList<>();
            Loop:
            while (true) {
                Temp:
                while (true) {

                    for (int i = 0; i < sep.size(); i++) {
                        if (sep.get(i) != null && sep.get(i)[0].indexOf(ctrlText.getText()) != -1) {
                            ctrlInd = i + 2;
                        }
                    }

                    for (int i = expInd; i < sep.size(); i++) {
                        if (sep.get(i)[0].equals("") | i == 0) {
                            expInd = i + 3;
                            break Temp;
                        }
                    }
                    break Loop;
                }
                System.out.println("HI1");
                while (true) {
                    if (ctrlInd != sep.size() && sep.get(expInd)[0].equals("Chemi")) {
                        sep.get(expInd)[5] = sep.get(expInd)[5].replace(",", "");
                        sep.get(ctrlInd)[5] = sep.get(ctrlInd)[5].replace(",", "");
                        comp.add(Double.parseDouble(sep.get(expInd)[5]) / Double.parseDouble(sep.get(ctrlInd)[5]));
                        expInd++;
                        ctrlInd++;
                    } else {
                        break;
                    }
                }
                comp.add(0d);
            }
            System.out.println("HI");
            saveT = new File(save.getAbsolutePath().substring(0, save.getAbsolutePath().length() - 4) + "txt");
            saveT.deleteOnExit();
            write = new PrintWriter(saveT);
            int indexT = 0;
            int indexV = 0;
            for (int i = indexT; i < store.size(); i++) {
                if (!sep.get(i)[0].equals("") | i == 0) {
                    write.println(store.get(i));
                } else {
                    indexT = i + 1;
                    write.println();
                    for (int q = indexV; q < comp.size(); q++) {
                        if (!comp.get(q).equals(0d) & indexV != 0) {
                            write.println(((q - indexV) % Double.parseDouble(incText.getText()) == 0 ? "ctrl" : "\t") + "\t\t\t\t\t" + Math.round(comp.get(q) * 100.0) / 100.0 + "\t" + (Math.round((comp.get(q) / comp.get((int) (q - q % indexV % Double.parseDouble(incText.getText())))) * 10000.0) / 100.0));
                        } else if (!comp.get(q).equals(0d)) {
                            write.println((q % Double.parseDouble(incText.getText()) == 0 ? "ctrl" : "\t") + "\t\t\t\t\t" + Math.round(comp.get(q) * 100.0) / 100.0 + "\t" + (Math.round((comp.get(q) / comp.get((int) (q - q % Double.parseDouble(incText.getText())))) * 10000.0) / 100.0));
                        } else {
                            indexV = q + 1;
                            write.println("\t");
                            break;
                        }
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (read != null) {
                read.close();
            }
            if (write != null) {
                write.close();
            }
        }
    }

    public static void print() {
//         Creating an instance of HSSFWorkbook.
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Create two sheets in the excel document and name it First Sheet and
        // Second Sheet.
        XSSFSheet firstSheet = workbook.createSheet("FIRST SHEET");

        // Manipulate the firs sheet by creating an HSSFRow which represent a
        // single row in excel sheet, the first row started from 0 index. After
        // the row is created we create a HSSFCell in this first cell of the row
        // and set the cell value with an instance of HSSFRichTextString
        // containing the words FIRST SHEET.
        XSSFRow rowA = firstSheet.createRow(0);
        XSSFCell cellA = rowA.createCell(0);
//        cellA.setCellValue(new HSSFRichTextString("FIRST SHEET"));

        // To write out the workbook into a file we need to create an output
        // stream where the workbook content will be written to.
        try (FileOutputStream fos = new FileOutputStream(new File(save.getAbsolutePath()))) {
            workbook.write(fos);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
        }
    }

    public static void convertCSV() throws IOException {
        // Sets up the Workbook and gets the 1st (0) sheet.
        print();
//        try (FileOutputStream fos = new FileOutputStream(new File("output.xls"))) {
//            HSSFWorkbook workbook = new HSSFWorkbook();
//            workbook.write(fos);
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
        File excelFile = new File(save.getAbsolutePath());
        InputStream input = new FileInputStream(excelFile);
        XSSFWorkbook workbook = new XSSFWorkbook(input);
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rowNo = 0;
        int columnNo = 0;

        // Loop through the files.
        Scanner scanner = new Scanner(new File(save.getAbsolutePath().substring(0, save.getAbsolutePath().length() - 4) + "txt"));
        while (scanner.hasNextLine()) {
            String line = scanner.nextLine();
            // Gets the row and if it doesn't exist it will create it.
            Row tempRow = sheet.getRow(rowNo);
            if (tempRow == null) {
                tempRow = sheet.createRow(rowNo);
            }

            Scanner lineScanner = new Scanner(line);
            lineScanner.useDelimiter("\t");
            // While there is more text to get it will loop.
            while (lineScanner.hasNext()) {
                // Gets the cell in that row and if it is null then it will be created.
                Cell tempCell = tempRow.getCell(columnNo, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String output = lineScanner.next();
                // Write the output to that cell.
                tempCell.setCellValue(new XSSFRichTextString(output));
                columnNo++;
            }
            lineScanner.close();
            // Resets the column count for the new row.
            columnNo = 0;
            rowNo++;
        }
        // Writes the file and closes everything.
        FileOutputStream out = new FileOutputStream(excelFile);
        workbook.write(out);
        input.close();
        out.close();
        scanner.close();
    }

    public static File convertExcel(File file) throws IOException {
        File newF = new File(file.getAbsolutePath().substring(0, file.getAbsolutePath().length() - 4) + "txt");
        try (PrintWriter write = new PrintWriter(newF)) {
            InputStream input = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(input);
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = firstSheet.iterator();

            while (iterator.hasNext()) {
                Row row = iterator.next();
                Iterator<Cell> iteratorC = row.cellIterator();
                while (iteratorC.hasNext()) {
                    Cell cell = iteratorC.next();
                    if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        break;
                    } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        write.print(cell.getStringCellValue() + "\t");
                    } else {
                        write.print(Double.toString(cell.getNumericCellValue()) + "\t");
                    }
                }
                write.print("\n");
            }
            System.out.println("FINAL");
            workbook.close();
        } catch (Exception e) {
            System.out.println("HI1234");
        } finally {
            if (write != null) {
                write.close();
            }
        }
        return newF;
    }
}
