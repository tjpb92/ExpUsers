package expusers;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
//import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Programme pour exporter les utilisateurs d'un site Web dans un fichier Excel
 *
 * @author Thierry Baribaud
 * @version Octobre 2016
 */
public class ExpUsers {

    private final static String filename = "users.xlsx";
    private static Object XHSSFCellStyle;

    /**
     * @param args arguments en ligne de commande
     */
    public static void main(String[] args) {

        FileOutputStream out;
        XSSFWorkbook classeur;
        XSSFSheet feuille;
        XSSFRow titre;
        XSSFCell cell;
        XSSFRow ligne = null;
        XSSFCellStyle blackMediumBorder = null;

        // Création d'un classeur Excel
        classeur = new XSSFWorkbook();
        feuille = classeur.createSheet("Utilisateurs");
        titre = feuille.createRow(0);
//        cell = titre.createCell((short) 0);
//        cell.setCellValue("Nom");

        // Style de cellule avec bordure noire
        blackMediumBorder = classeur.createCellStyle();
        blackMediumBorder.setBorderBottom(BorderStyle.THIN);
        blackMediumBorder.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        blackMediumBorder.setBorderLeft(BorderStyle.THIN);
        blackMediumBorder.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        blackMediumBorder.setBorderRight(BorderStyle.THIN);
        blackMediumBorder.setRightBorderColor(IndexedColors.BLACK.getIndex());
        blackMediumBorder.setBorderTop(BorderStyle.THIN);
        blackMediumBorder.setTopBorderColor(IndexedColors.BLACK.getIndex());

        // Intialisitation du titre
        cell = titre.createCell((short) 0);
        cell.setCellStyle(blackMediumBorder);
        cell.setCellValue("Nom");
        cell = titre.createCell((short) 1);
        cell.setCellStyle(blackMediumBorder);
        cell.setCellValue("Prénom");
        cell = titre.createCell((short) 2);
        cell.setCellStyle(blackMediumBorder);
        cell.setCellValue("Niveau");
        cell = titre.createCell((short) 3);
        cell.setCellStyle(blackMediumBorder);
        cell.setCellValue("Mail");
        cell = titre.createCell((short) 4);
        cell.setCellStyle(blackMediumBorder);
        cell.setCellValue("Etat");
        cell = titre.createCell((short) 4);
        cell.setCellStyle(blackMediumBorder);
        cell.setCellValue("Société");

        // Initialisation tableau
        for (int i = 0; i < 10; i++) {
            ligne = feuille.createRow(i + 1);
            for (int j = 0; j < 5; j++) {
                cell = ligne.createCell((short) j);
                cell.setCellValue(i * j);
                cell.setCellStyle(blackMediumBorder);
            }
        }

        // Enregistrement du classeur dans un fichier
        try {
            out = new FileOutputStream(new File(filename));
            classeur.write(out);
            out.close();
            System.out.println("Fichier Excel " + filename + " créé");
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExpUsers.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExpUsers.class.getName()).log(Level.SEVERE, null, ex);
        }

    }
}
