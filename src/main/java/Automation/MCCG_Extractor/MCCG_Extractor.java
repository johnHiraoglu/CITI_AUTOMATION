package Automation.MCCG_Extractor;

import Utils.ExcelUtil;

import javax.swing.*;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;

public class MCCG_Extractor {
    public static void main(String[] args) {

        //StringBuilder for saving output
        StringBuilder st = new StringBuilder();

        //Column Number of MCC to be extracted
        int columnNumber = 0;

        //File Path
        String path = "";

        //JFile Chooser. Choosing the file.
        JFileChooser fileChooser = new JFileChooser();

        //Enter your desktop path with in " "
        fileChooser.setCurrentDirectory(new File("C:\\Users\\John\\Desktop"));
        int result = fileChooser.showOpenDialog(fileChooser);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            path = selectedFile.getAbsolutePath();
        }

        //Pulling in Excel File & Sheet
        ExcelUtil excelUtil = new ExcelUtil(path, "MCC US_CANADA");

        //Capturing the Column Number
        String iconArray[] = {"D", "E", "F", "G", "H"};
        Object selection = JOptionPane.showOptionDialog(null, "Select the MCCG Column", "MCCG Extractor",
                JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, iconArray, iconArray[0]);

        if (selection.equals(0)) {
            columnNumber = 3;
        } else if (selection.equals(1)) {
            columnNumber = 4;
        } else if (selection.equals(2)) {
            columnNumber = 5;
        } else if (selection.equals(3)) {
            columnNumber = 6;
        } else if (selection.equals(4)) {
            columnNumber = 7;
        } else {
            System.out.println("Invalid Option Chosen");
        }


        //MCCG Action
        String iconArray1[] = {"Include", "Exclude"};
        String action = "";
        Object selection1 = JOptionPane.showOptionDialog(null, "Select the MCCG Action", "MCCG Extractor",
                JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, null, iconArray1, iconArray1[0]);

        if (selection1.equals(0)) {
            action = "I";
        } else if (selection.equals(1)) {
            action = "E";
        } else {
            System.out.println("Invalid Option Chosen");
        }

        try {
            File fout = new File("MCCG output.csv");
            FileOutputStream fos = new FileOutputStream(fout);

            BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));


            //Printing out the Template name
            st.append("\n\nTemplate Name = " + excelUtil.getCellData(12, columnNumber));
            bw.write("Template Name," + excelUtil.getCellData(12, columnNumber));
            bw.newLine();
            st.append(System.lineSeparator());

            //Printing MCCG Action
            if (action.equals("I")) {
                st.append("MCCG Action = Include");
                bw.write("MCCG Action,Include");
                bw.newLine();

            } else
                st.append("MCCG Action = Exclude");
            bw.write("MCCG Action,Exclude");
            bw.newLine();
            st.append(System.lineSeparator());

            //Looping to get All MCCs
            for (int i = 19; i < excelUtil.rowCount(); i++) {
                if (action.equals("I") && excelUtil.getCellData(i, columnNumber).equals("Unblock")) {
                    st.append(excelUtil.getCellData(i, 0).substring(0, 4) + " = " + excelUtil.getCellData(i, columnNumber));
                    st.append(System.lineSeparator());
                    bw.write(excelUtil.getCellData(i, 0).substring(0, 4) + "," + excelUtil.getCellData(i, columnNumber));
                    bw.newLine();
                }
                if (action.equals("E") && excelUtil.getCellData(i, columnNumber).equals("Block")) {
                    st.append(excelUtil.getCellData(i, 0).substring(0, 4) + " = " + excelUtil.getCellData(i, columnNumber));
                    st.append(System.lineSeparator());
                    bw.write(excelUtil.getCellData(i, 0).substring(0, 4) + "," + excelUtil.getCellData(i, columnNumber));
                    bw.newLine();
                }
            }

            bw.close();
        } catch (Exception e) {
            System.out.println("Problem");

        }

        //Printing the output
        System.out.println(st);
    }
}

