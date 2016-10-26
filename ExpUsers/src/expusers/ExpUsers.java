package expusers;

import bkgpi2a.User;
import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;
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
import org.bson.Document;

/**
 * Programme pour exporter les utilisateurs d'un site Web dans un fichier Excel
 *
 * @author Thierry Baribaud
 * @version Octobre 2016
 */
public class ExpUsers {

    private final static String filename = "users.xlsx";
//    private static Object XHSSFCellStyle;

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
        XSSFRow ligne = null;
        XSSFCellStyle blackMediumBorder = null;

        MongoDatabase MyDatabase;

        MongoClient MyMongoClient = new MongoClient(HOST, PORT);

        System.out.println("Liste des bases de données :");
        for (String MyDbNname : MyMongoClient.listDatabaseNames()) {
            System.out.println("  " + MyDbNname);
            MyDatabase = MyMongoClient.getDatabase(MyDbNname);
            System.out.println("Liste des collections de " + MyDbNname + " :");
            for (String MyCollectionName : MyDatabase.listCollectionNames()) {
                System.out.println("  " + MyCollectionName);
            }
        }

        MyDatabase = MyMongoClient.getDatabase("extranet");
        MongoCollection<Document> MyCollection = MyDatabase.getCollection("users");
        System.out.println(MyCollection.count() + " utilisateurs");

        Document MyDocument = MyCollection.find().first();
        System.out.println(MyDocument.toJson());

        
//      Création d'un classeur Excel
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

//        MongoCursor<User> cursor = MyCollection.find(User.class).iterator();
//        try {
//            while (cursor.hasNext()) {
//                System.out.println(cursor.next());
//            }
//        } finally {
//            cursor.close();
//        }
        Document user;
        Document company;
        MongoCursor<Document> MyCursor = MyCollection.find().iterator();
        int n=0;
        try {
            while (MyCursor.hasNext()) {
                user = MyCursor.next();
//                company = user.get(company);
                System.out.println("lastName:" + user.getString("lastName") + 
                        ", firstName:" + user.getString("firstName") +
//                        ", userType:" + user.getString("userType") + 
//                        ", status:" + user.getBoolean("isActive") +
//                        ", company:" + (Document) (user.get("company")).getString("label")) + 
                        ", login:" + user.getString("login"));
//                System.out.println(MyCursor.next().toJson());
                n++;
                ligne = feuille.createRow(n);
    
                cell = ligne.createCell(0);
                cell.setCellValue(user.getString("lastName"));
                cell.setCellStyle(blackMediumBorder);

                cell = ligne.createCell(1);
                cell.setCellValue(user.getString("firstName"));
                cell.setCellStyle(blackMediumBorder);

                cell = ligne.createCell(2);
                cell.setCellValue(user.getString("userType"));
                cell.setCellStyle(blackMediumBorder);

                cell = ligne.createCell(3);
                cell.setCellValue(user.getString("login"));
                cell.setCellStyle(blackMediumBorder);
            }
        } finally {
            MyCursor.close();
        }
        
//         Initialisation tableau
//        for (int i = 0; i < 10; i++) {
//            ligne = feuille.createRow(i + 1);
//            for (int j = 0; j < 5; j++) {
//                cell = ligne.createCell((short) j);
//                cell.setCellValue(i * j);
//                cell.setCellStyle(blackMediumBorder);
//            }
//        }

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
