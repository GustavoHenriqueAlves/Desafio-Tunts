import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.Collections;
import java.util.List;

public class programm {

    private static Sheets sheetsService;
    private static final String SPREADSHEET_ID = "1leCaLVizq5d7a9ZKNyDP23Tx0SR7L6NItSvSp1nf-dE";

    private static Credential authorize() throws IOException, GeneralSecurityException{

        System.out.println("Authorizing service");

        InputStream in = programm.class.getResourceAsStream( "/credentials.json");
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(
                JacksonFactory.getDefaultInstance(), new InputStreamReader(in)
        );

        List<String> scopes = Collections.singletonList(SheetsScopes.SPREADSHEETS);

        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(),
                clientSecrets, scopes)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File("tokens")))
                .setAccessType("offline")
                .build();

        return new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver())
                .authorize("user");
    }

    public static Sheets getSheetsService() throws IOException, GeneralSecurityException{

        System.out.println("Getting Sheets Service...");

        Credential credential = authorize();

        String APPLICATION_NAME = "Tunts Challenge";
        return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(),
                    JacksonFactory.getDefaultInstance(), credential)
                    .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public static ValueRange generateValueRange(String range) throws IOException {

        System.out.println("Generating Value Range...");

        return sheetsService.spreadsheets().values()
                .get(SPREADSHEET_ID, range)
                .execute();
    }

    public static int[] generateArray(List<List<Object>> list){

        System.out.println("Generating Arrays...");

        int i = 0;
        int[] array = new int[24];

        for( List row : list){

            array[i]= Integer.parseInt((row.get(0).toString()));
            i++;
        }

        return array;
    }

    public static void updateSheet(String situation, int naf, int a) throws IOException {

        System.out.println("Updating Sheet...");

        String situationRange = "G" + a;
        String nfaRange = "H" + a;

        ValueRange body = new ValueRange()
                .setValues(Collections.singletonList(Collections.singletonList(situation)));

        UpdateValuesResponse result = sheetsService.spreadsheets().values()
                .update(SPREADSHEET_ID, situationRange, body)
                .setValueInputOption("RAW")
                .execute();

        body = new ValueRange()
                .setValues(Collections.singletonList(Collections.singletonList(naf)));

        result = sheetsService.spreadsheets().values()
                .update(SPREADSHEET_ID, nfaRange, body)
                .setValueInputOption("RAW")
                .execute();
    }

    public static void main(String[] args) throws IOException, GeneralSecurityException{

        sheetsService = getSheetsService();
        String absenceRange = "C4:C";
        String p1Range = "D4:D";
        String p2Range = "E4:E";
        String p3Range = "F4:F";

        ValueRange p1Response = generateValueRange(p1Range);
        ValueRange p2Response = generateValueRange(p2Range);
        ValueRange p3Response = generateValueRange(p3Range);
        ValueRange absenceResponse = generateValueRange(absenceRange);

        List<List<Object>> p1List = p1Response.getValues();
        List<List<Object>> p2List = p2Response.getValues();
        List<List<Object>> p3List = p3Response.getValues();
        List<List<Object>> absenceList = absenceResponse.getValues();

        int[] p1Array = generateArray(p1List);
        int[] p2Array = generateArray(p2List);
        int[] p3Array = generateArray(p3List);
        int[] absenceArray = generateArray(absenceList);

        int a = 4;
        int b = 0;

        for (int i : absenceArray){
            
            int mnoa = 15;

            if (i > mnoa){

                updateSheet("Reprovado por falta", 0, a);
            }else{

                int grade = p1Array[b] + p2Array[b] + p3Array[b];
                int m = Math.round(grade / 3);

                if (m < 50){

                    updateSheet("Reprovado por nota", 0, a);
                }else if (m < 70){

                    int naf = Math.round(100 - m);
                    updateSheet("Exame Final", naf, a);
                }else {

                    updateSheet("Aprovado", 0, a);
                }
            }

            a++;
            b++;
        }

        System.out.println("Ending the program...");
        System.exit(0);
    }
}