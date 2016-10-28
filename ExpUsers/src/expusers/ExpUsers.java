package expusers;

import bkgpi2a.CallCenterUser;
import bkgpi2a.ClientAccountManager;
import bkgpi2a.Executive;
import bkgpi2a.SuperUser;
import bkgpi2a.User;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.mongodb.BasicDBObject;
import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PaperSize;
import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;

/**
 * Programme pour exporter les utilisateurs d'un site Web dans un fichier Excel
 *
 * @author Thierry Baribaud
 * @version Octobre 2016
 */
public class ExpUsers {

    private final static String path = "c:\\temp";
    
    private final static String filename = "users.xlsx";

    private final static String HOST = "192.168.0.17";
    private final static int PORT = 27017;

    /**
     * @param args arguments en ligne de commande
     */
    public static void main(String[] args) {

        FileOutputStream out;
        XSSFWorkbook classeur;
        XSSFSheet feuille;
        XSSFRow titre;
        XSSFCell cell;
        XSSFRow ligne;
        XSSFCellStyle cellStyle;
        XSSFCellStyle titleStyle;
        ObjectMapper objectMapper;
        User user;
        MongoDatabase mongoDatabase;
        MongoClient MyMongoClient;

        objectMapper = new ObjectMapper();

        MyMongoClient = new MongoClient(HOST, PORT);
        mongoDatabase = MyMongoClient.getDatabase("extranet");

        MongoCollection<Document> MyCollection = mongoDatabase.getCollection("users");
        System.out.println(MyCollection.count() + " utilisateurs");

//      Création d'un classeur Excel
        classeur = new XSSFWorkbook();
        feuille = classeur.createSheet("Utilisateurs");
        
        // Style de cellule avec bordure noire
        cellStyle = classeur.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

        // Style pour le titre
        titleStyle = (XSSFCellStyle) cellStyle.clone();
        titleStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        titleStyle.setFillPattern(FillPatternType.LESS_DOTS);
//        titleStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());

        // Ligne de titre
        titre = feuille.createRow(0);
        cell = titre.createCell((short) 0);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Nom");
        cell = titre.createCell((short) 1);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Prénom");
        cell = titre.createCell((short) 2);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Niveau");
        cell = titre.createCell((short) 3);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Mail");
        cell = titre.createCell((short) 4);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Etat");
        cell = titre.createCell((short) 5);
        cell.setCellStyle(titleStyle);
        cell.setCellValue("Société");

        // Lit les ustilisateurs classés par nom et prénom
        MongoCursor<Document> MyCursor
                = MyCollection.find().sort(new BasicDBObject("lastName", 1).append("firstName", 1)).iterator();
        int n = 0;
        try {
            while (MyCursor.hasNext()) {
                user = objectMapper.readValue(MyCursor.next().toJson(), User.class);
                System.out.println( n 
                        + " lastName:" + user.getLastName()
                        + ", firstName:" + user.getFirstName()
                        + ", status:" + user.getIsActive()
                        + ", login:" + user.getLogin()
                        + ", class:" + user.getClass().getSimpleName());
                n++;
                ligne = feuille.createRow(n);

                cell = ligne.createCell(0);
                cell.setCellValue(user.getLastName());
                cell.setCellStyle(cellStyle);

                cell = ligne.createCell(1);
                cell.setCellValue(user.getFirstName());
                cell.setCellStyle(cellStyle);

                if (user instanceof Executive) {
                    cell = ligne.createCell(2);
                    cell.setCellValue(((Executive) user).getClass().getSimpleName());
                    cell.setCellStyle(cellStyle);

                    cell = ligne.createCell(5);
                    cell.setCellValue(((Executive) user).getCompany().getLabel());
                    cell.setCellStyle(cellStyle);
                } else if (user instanceof CallCenterUser) {
                    cell = ligne.createCell(2);
                    cell.setCellValue(((CallCenterUser) user).getClass().getSimpleName());
                    cell.setCellStyle(cellStyle);

                    cell = ligne.createCell(5);
                    cell.setCellValue(((CallCenterUser) user).getCompany().getLabel());
                    cell.setCellStyle(cellStyle);
                } else if (user instanceof ClientAccountManager) {
                    cell = ligne.createCell(2);
                    cell.setCellValue(((ClientAccountManager) user).getClass().getSimpleName());
                    cell.setCellStyle(cellStyle);

                    cell = ligne.createCell(5);
                    cell.setCellValue(((ClientAccountManager) user).getCompany().getLabel());
                    cell.setCellStyle(cellStyle);
                } else if (user instanceof SuperUser) {
                    cell = ligne.createCell(2);
                    cell.setCellValue(((SuperUser) user).getClass().getSimpleName());
                    cell.setCellStyle(cellStyle);

                    cell = ligne.createCell(5);
                    cell.setCellStyle(cellStyle);
                }

                cell = ligne.createCell(3);
                cell.setCellValue(user.getLogin());
                cell.setCellStyle(cellStyle);

                cell = ligne.createCell(4);
                if (user.getIsActive()) {
                    cell.setCellValue("Actif");
                } else {
                    cell.setCellValue("Inactif");
                }
                cell.setCellStyle(cellStyle);
            }

            // Ajustement automatique de la largeur des colonnes
            for (int k = 0; k < 6; k++) {
                feuille.autoSizeColumn(k);
            }

            // Format A4 en sortie
            feuille.getPrintSetup().setPaperSize(PaperSize.A4_PAPER);

            // Orientation paysage
            feuille.getPrintSetup().setLandscape(true);

            // Ajustement à une page en largeur
            feuille.setFitToPage(true);
            feuille.getPrintSetup().setFitWidth((short) 1);
            feuille.getPrintSetup().setFitHeight((short) 0);

            // En-tête et pied de page
            Header header  = feuille.getHeader();
            header.setLeft("Liste des utilisateurs Extranet Anstel");
            header.setRight("&F");
            
            Footer footer = feuille.getFooter();
            footer.setLeft("Documentation confidentielle Anstel");
            footer.setCenter("Page &P / &N");
            footer.setRight("&D");

            // Ligne à répéter en haut de page
            feuille.setRepeatingRows(CellRangeAddress.valueOf("1:1"));
            
        } catch (IOException ex) {
            Logger.getLogger(ExpUsers.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            MyCursor.close();
        }

        // Enregistrement du classeur dans un fichier
        try {
            out = new FileOutputStream(new File(path + "\\" + filename));
            classeur.write(out);
            out.close();
            System.out.println("Fichier Excel " + filename + " créé dans " + path);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExpUsers.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExpUsers.class.getName()).log(Level.SEVERE, null, ex);
        }

    }
}
