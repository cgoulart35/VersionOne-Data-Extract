package com.versiononedataextract;

import org.apache.http.config.Registry;
import org.apache.http.config.RegistryBuilder;
import org.apache.http.conn.socket.ConnectionSocketFactory;
import org.apache.http.conn.socket.PlainConnectionSocketFactory;
import org.apache.http.conn.ssl.NoopHostnameVerifier;
import org.apache.http.conn.ssl.SSLConnectionSocketFactory;
import org.apache.http.conn.ssl.TrustStrategy;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.impl.conn.BasicHttpClientConnectionManager;
import org.apache.http.ssl.SSLContexts;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.commons.lang3.StringUtils;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.http.*;
import org.springframework.http.client.HttpComponentsClientHttpRequestFactory;
import org.springframework.web.client.RestTemplate;

import javax.net.ssl.SSLContext;

import java.io.*;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.KeyManagementException;
import java.security.KeyStoreException;
import java.security.NoSuchAlgorithmException;
import java.util.*;
import java.util.List;

public class GenerateExcelDataPart2 {

    private static final String GLOBAL_DIRECTORY_PATH = System.getProperty("user.home") + File.separatorChar + "Documents";

    private static String company;

    private static String bearerAuthToken;

    private static void setBearerAuthToken(final String token) {
        bearerAuthToken = token;
    }

    private static void setCompany(final String companyToken) {
        company = companyToken;
    }

    private final static class ExcelWorkbook {
        private final XSSFWorkbook wb = new XSSFWorkbook();

        private List<String> projectSheetColumns = new ArrayList<>();
        private int projectSheetRowIndex = 1;

        private List<String> childProjectSheetColumns = new ArrayList<>();
        private int childProjectSheetRowIndex = 1;

        private List<String> epicSheetColumns = new ArrayList<>();
        private int epicSheetRowIndex = 1;

        private List<String> storyDefectSheetColumns = new ArrayList<>();
        private int storyDefectSheetRowIndex = 1;

        private List<String> taskSheetColumns = new ArrayList<>();
        private int taskSheetRowIndex = 1;
    }

    enum Sheet_Type {
        Parent_Project_Sheet,
        Child_Project_Sheet,
        Epic_Sheet,
        Story_Defect_Sheet,
        Task_Sheet
    }

    private static ResponseEntity<String> getResponse(final String URL) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException {
        final TrustStrategy acceptingTrustStrategy = (cert, authType) -> true;
        final SSLContext sslContext = SSLContexts.custom().loadTrustMaterial(null, acceptingTrustStrategy).build();
        final SSLConnectionSocketFactory sslsf = new SSLConnectionSocketFactory(sslContext, NoopHostnameVerifier.INSTANCE);
        final Registry<ConnectionSocketFactory> socketFactoryRegistry = RegistryBuilder.<ConnectionSocketFactory>create()
                .register("https", sslsf)
                .register("http", new PlainConnectionSocketFactory())
                .build();
        final BasicHttpClientConnectionManager connectionManager = new BasicHttpClientConnectionManager(socketFactoryRegistry);
        final CloseableHttpClient httpClient = HttpClients.custom()
                .setSSLSocketFactory(sslsf)
                .setConnectionManager(connectionManager)
                .build();
        final HttpComponentsClientHttpRequestFactory requestFactory = new HttpComponentsClientHttpRequestFactory(httpClient);
        final RestTemplate restTemplate = new RestTemplate(requestFactory);
        final HttpHeaders headers= new HttpHeaders();
        headers.setBearerAuth(bearerAuthToken);
        final List<MediaType> accept = new ArrayList<MediaType>();
        accept.add(MediaType.APPLICATION_JSON);
        headers.setAccept(accept);
        final HttpEntity<String> entity = new HttpEntity<String>(headers);
        final ResponseEntity<String> result = restTemplate.exchange(URL, HttpMethod.GET, entity, String.class);
        if (result.getStatusCode() != HttpStatus.OK) {
            System.out.println("Response did not have status 200/OK." + "\n" + "URL called: " + URL + "\n" + "Response: " + result);
            throw new Exception ("Response did not have status 200/OK.");
        }
        return result;
    }

    enum Project_URL_Type {
        My_Project_URL,
        Base_URL,
        CreateDate_URL,
        CreatedBy_URL,
        ClosedDate_URL,
        ClosedBy_URL,
        DoneDate_URL,
        AcceptanceCriteria_URL,
        Done_Hours_URL,
        Done_Hours_List_URL,
        My_Child_Project_List_URL,
        Epic_List_URL,
        Story_List_URL,
        Story_List_In_Scope,
        Defect_List_URL,
        Defect_List_In_Scope,
        Task_List_URL,
        Actuals_List_URL
    }

    private static String getURL(final Project_URL_Type project_URL_Type, final String token) {
        switch (project_URL_Type) {
            case My_Project_URL:
                return "https://www5.v1host.com/" + company + "/rest-1.v1/Data/Scope/" + token;
            case Base_URL:
                return "https://www5.v1host.com" + token;
            case CreateDate_URL:
                return "https://www5.v1host.com" + token + "/CreateDate";
            case CreatedBy_URL:
                return "https://www5.v1host.com" + token + "/CreatedBy";
            case ClosedDate_URL:
                return "https://www5.v1host.com" + token + "/ClosedDate";
            case ClosedBy_URL:
                return "https://www5.v1host.com" + token + "/ClosedBy";
            case DoneDate_URL:
                return "https://www5.v1host.com" + token + "/DoneDate";
            case AcceptanceCriteria_URL:
                return "https://www5.v1host.com" + token + "/Custom_AcceptanceCriteria2";
            case Done_Hours_URL:
                return "https://www5.v1host.com" + token + "?sel=Name,Actuals.Value.@Sum";
            case Done_Hours_List_URL:
                return "https://www5.v1host.com/" + company + "/rest-1.v1/Data/Task?where=Parent='" + token + "'&sel=Name,Actuals.Value.@Sum";
            case My_Child_Project_List_URL:
                return "https://www5.v1host.com/" + company + "/rest-1.v1/Data/Scope?where=Parent='Scope:" + token + "'";
            case Epic_List_URL:
                return "https://www5.v1host.com/" + company + "/rest-1.v1/Data/Epic?where=Scope='" + token + "'";
            case Story_List_URL:
                return "https://www5.v1host.com/" + company + "/rest-1.v1/Data/Story?where=Super='" + token + "'";
            case Story_List_In_Scope:
                return "https://www5.v1host.com/" + company + "/rest-1.v1/Data/Story?where=Scope='" + token + "'";
            case Defect_List_URL:
                return "https://www5.v1host.com/" + company + "/rest-1.v1/Data/Defect?where=Super='" + token + "'";
            case Defect_List_In_Scope:
                return "https://www5.v1host.com/" + company + "/rest-1.v1/Data/Defect?where=Scope='" + token + "'";
            case Task_List_URL:
                return "https://www5.v1host.com/" + company + "/rest-1.v1/Data/Task?where=Parent='" + token + "'";
            case Actuals_List_URL:
                return "https://www5.v1host.com" + token + "?sel=Actuals";
            default:
                return null;
        }
    }

    private static void createNextDirectory(final String targetDirectory) throws IOException {
        final Path path = Paths.get(targetDirectory);
        if (!Files.exists(path)) {
            Files.createDirectory(path);
            System.out.println("Directory created successfully: " + targetDirectory);
        } else {
            System.out.println("Directory already exists: " + targetDirectory);
        }
    }

    private static String makeAcceptableName(final String name) {
        return StringUtils.left(name.replaceAll("[^a-zA-Z0-9\\-]", "_"), 127);
    }

    private static String extractInformationOfParentProject(final ExcelWorkbook myExcelWorkbook, final String projectIDNumber) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.My_Project_URL, projectIDNumber));  // GET PROJECT JSON RESPONSE (URL IS HARD CODED)
        final JSONObject outerObject = new JSONObject(response.getBody());
        final String projectScopeToken = outerObject.getString("href");
        final JSONObject attributeObject = outerObject.getJSONObject("Attributes");
        final JSONObject nameObject = attributeObject.getJSONObject("Name");
        final String name = nameObject.optString("value");
        final int rowNum = parseJSONObjectBodyToMapForExcel(myExcelWorkbook, outerObject, Sheet_Type.Parent_Project_Sheet);
        extractCreateData(myExcelWorkbook, projectScopeToken, Sheet_Type.Parent_Project_Sheet, rowNum);                                  // GET DATA ON CREATED BY AND DATE
        extractClosedData(myExcelWorkbook, projectScopeToken, Sheet_Type.Parent_Project_Sheet, rowNum);                                  // GET DATA ON CLOSED BY AND DATE
        double totalEstimate = extractInformationOfStoriesAndDefects(myExcelWorkbook, true, "Scope:" + projectIDNumber);
        totalEstimate = totalEstimate + extractInformationOfChildProjects(myExcelWorkbook, projectIDNumber);                             // GET DATA ON CHILD PROJECTS IN PROJECT (URL IS HARD CODED)
        extractEstimatePTS(myExcelWorkbook, totalEstimate, true, rowNum, Sheet_Type.Parent_Project_Sheet);
        return name;
    }

    private static double extractInformationOfChildProjects(final ExcelWorkbook myExcelWorkbook, final String projectIDNumber) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.My_Child_Project_List_URL, projectIDNumber));    // GET JSON RESPONSE LIST OF CHILD PROJECTS (URL IS HARD CODED)
        final JSONObject outerObjectChildren = new JSONObject(response.getBody());
        final JSONArray assetsJsonArrayChildren = outerObjectChildren.getJSONArray("Assets");
        double totalEstimate = 0;
        int childProjectIndex = 0;
        while (childProjectIndex < assetsJsonArrayChildren.length()) {                                                              // CREATE CHILD PROJECT DIRECTORIES + THEIR JSON FILES
            final JSONObject childProject = assetsJsonArrayChildren.getJSONObject(childProjectIndex);                       // read next child project in list
            totalEstimate = totalEstimate + extractInformationOfChildProject(myExcelWorkbook, childProject, Sheet_Type.Child_Project_Sheet);
            childProjectIndex++;
        }
        return totalEstimate;
    }

    private static double extractInformationOfChildProject(final ExcelWorkbook myExcelWorkbook, final JSONObject childProject, Sheet_Type sheet_type) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String childProjectScopeToken = childProject.getString("href");                                              // get child scope token
        final String childProjectID = childProject.getString("id");                                                        // get child id
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, childProjectScopeToken));         // get child project json response
        final JSONObject outerObject = new JSONObject(response.getBody());
        final int rowNum = parseJSONObjectBodyToMapForExcel(myExcelWorkbook, outerObject, sheet_type);
        extractCreateData(myExcelWorkbook, childProjectScopeToken, sheet_type, rowNum);                      // GET DATA ON CREATED BY AND DATE
        extractClosedData(myExcelWorkbook, childProjectScopeToken, sheet_type, rowNum);                      // GET DATA ON CLOSED BY AND DATE
        double totalEstimate = extractInformationOfStoriesAndDefects(myExcelWorkbook, true, childProjectID);
        totalEstimate = totalEstimate + extractInformationOfEpics(myExcelWorkbook, childProjectID);             // GET DATA ON EPICS IN CHILD PROJECT
        extractEstimatePTS(myExcelWorkbook, totalEstimate, true, rowNum, sheet_type);
        return totalEstimate;
    }

    private static double extractInformationOfEpics(final ExcelWorkbook myExcelWorkbook, final String childProjectID) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Epic_List_URL, childProjectID));                    // GET JSON RESPONSE LIST OF EPICS IN CHILD PROJECT
        final JSONObject outerObjectEpics = new JSONObject(response.getBody());
        final JSONArray assetsJsonArrayEpics = outerObjectEpics.getJSONArray("Assets");
        double totalEstimate = 0;
        int epicIndex = 0;
        while (epicIndex < assetsJsonArrayEpics.length()) {                                                                             // CREATE EPIC PROJECT DIRECTORIES + THEIR JSON FILES
            final JSONObject epic = assetsJsonArrayEpics.getJSONObject(epicIndex);                                                            // read next epic in list
            final JSONObject epicAttributeObject = epic.getJSONObject("Attributes");
            final JSONObject epicSuperObject = epicAttributeObject.getJSONObject("Super");
            final JSONObject epicSuperValueObject = epicSuperObject.optJSONObject("value");
            if (epicSuperValueObject == null) {                                                                                              //find all Epics in a child project that have no super epic (could be an epic or sub-epic)
                totalEstimate = totalEstimate + extractInformationOfEpic(myExcelWorkbook, epic, childProjectID);
            }
            epicIndex++;
        }
        return totalEstimate;
    }

    private static double extractInformationOfEpic(final ExcelWorkbook myExcelWorkbook, final JSONObject epic, final String childProjectIDForAssertion) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        double totalEstimate = 0;
        final JSONObject epicAttributeObject = epic.getJSONObject("Attributes");
        final String epicScopeToken = epic.getString("href");                                                                              // get epic scope token
        final String epicID = epic.getString("id");                                                                                        //get epic id
        final ResponseEntity<String> response1 = getResponse(getURL(Project_URL_Type.Base_URL, epicScopeToken));                                // get epic json response
        final JSONObject outerObject = new JSONObject(response1.getBody());
        final int rowNum = parseJSONObjectBodyToMapForExcel(myExcelWorkbook, outerObject, Sheet_Type.Epic_Sheet);
        extractCreateData(myExcelWorkbook, epicScopeToken, Sheet_Type.Epic_Sheet, rowNum);                                                          // GET DATA ON CREATED BY AND DATE
        extractClosedData(myExcelWorkbook, epicScopeToken, Sheet_Type.Epic_Sheet, rowNum);                                                          // GET DATA ON CLOSED BY AND DATE
        extractDoneDate(myExcelWorkbook, epicScopeToken, Sheet_Type.Epic_Sheet, rowNum);                                                            // GET DATA ON DONE DATE
        String childProjectID = childProjectIDForAssertion;
        if(childProjectIDForAssertion.equals("notAnID")) {
            final JSONObject epicSecurityScopeObject = epicAttributeObject.getJSONObject("SecurityScope");                                      // get epic security scope
            final JSONObject epicSecurityScopeValueObject = epicSecurityScopeObject.getJSONObject("value");
            childProjectID = epicSecurityScopeValueObject.getString("idref");
        }
        totalEstimate = totalEstimate + extractInformationOfStoriesAndDefects(myExcelWorkbook, false, epicID);                                            // GET DATA ON DEFECTS AND STORIES IN SUB-EPIC
        final ResponseEntity<String> response2 = getResponse(getURL(Project_URL_Type.Epic_List_URL, childProjectID));             // GET JSON RESPONSE LIST OF EPICS IN CHILD PROJECT
        final JSONObject outerObjectEpics = new JSONObject(response2.getBody());
        final JSONArray assetsJsonArrayEpics = outerObjectEpics.getJSONArray("Assets");
        totalEstimate = totalEstimate + extractInformationOfSubEpics(myExcelWorkbook, assetsJsonArrayEpics, childProjectID, epicID);               // GET DATA ON SUB-EPICS IN EPIC
        extractEstimatePTS(myExcelWorkbook, totalEstimate, false, rowNum, Sheet_Type.Epic_Sheet);
        return totalEstimate;
    }

    private static double extractInformationOfSubEpics(final ExcelWorkbook myExcelWorkbook, final JSONArray assetsJsonArrayEpics, final String childProjectID, final String epicID ) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        double totalEstimate = 0;
        int subEpicIndex = 0;
        while(subEpicIndex < assetsJsonArrayEpics.length()) {                                                                               // CREATE SUB-EPIC/CHILD EPIC PROJECT DIRECTORIES + THEIR JSON FILES
            final JSONObject subEpic = assetsJsonArrayEpics.getJSONObject(subEpicIndex);                                                                     // read next epic in list
            final JSONObject subEpicAttributeObject = subEpic.getJSONObject("Attributes");
            final JSONObject subEpicSuperObject = subEpicAttributeObject.getJSONObject("Super");                                                             //find all Epics in a child project that are sub epics/epics of the current epic
            final JSONObject subEpicSuperValueObject = subEpicSuperObject.optJSONObject("value");
            if (subEpicSuperValueObject != null) {
                final String subEpicSuperValue = subEpicSuperValueObject.getString("idref");
                if (subEpicSuperValue.equals(epicID)) {
                    final String subEpicScopeToken = subEpic.getString("href");                                                                         // get sub-epic scope token
                    final String subEpicID = subEpic.getString("id");                                                                                   //get sub-epic id`
                    final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, subEpicScopeToken));                               // get sub-epic json response
                    final JSONObject outerObject = new JSONObject(response.getBody());
                    final int rowNum = parseJSONObjectBodyToMapForExcel(myExcelWorkbook, outerObject, Sheet_Type.Epic_Sheet);
                    extractCreateData(myExcelWorkbook, subEpicScopeToken, Sheet_Type.Epic_Sheet, rowNum);                                                           // GET DATA ON CREATED BY AND DATE
                    extractClosedData(myExcelWorkbook, subEpicScopeToken, Sheet_Type.Epic_Sheet, rowNum);                                                           // GET DATA ON CLOSED BY AND DATE
                    extractDoneDate(myExcelWorkbook, subEpicScopeToken, Sheet_Type.Epic_Sheet, rowNum);                                                             // GET DATA ON DONE DATE
                    double subTotalEstimate = extractInformationOfStoriesAndDefects(myExcelWorkbook, false, subEpicID) + extractInformationOfSubEpics(myExcelWorkbook, assetsJsonArrayEpics, childProjectID, subEpicID);;
                    extractEstimatePTS(myExcelWorkbook, subTotalEstimate, false, rowNum, Sheet_Type.Epic_Sheet);
                    totalEstimate = totalEstimate + subTotalEstimate;                                                                       // GET DATA ON DEFECTS AND STORIES IN SUB-EPIC AND DATA ON SUB-EPICS OF CURRENT CHILD EPIC
                }
            }
            subEpicIndex++;
        }
        return totalEstimate;
    }

    private static double extractInformationOfStoriesAndDefects(final ExcelWorkbook myExcelWorkbook, final boolean isProject, final String id) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final Project_URL_Type storyURL;
        final Project_URL_Type defectURL;
        if (isProject) {
            storyURL = Project_URL_Type.Story_List_In_Scope;
            defectURL = Project_URL_Type.Defect_List_In_Scope;
        } else {
            storyURL = Project_URL_Type.Story_List_URL;
            defectURL = Project_URL_Type.Defect_List_URL;
        }
        final ResponseEntity<String> response1 = getResponse(getURL(storyURL, id));                      // GET JSON RESPONSE LIST OF STORIES/DEFECTS IN EPIC
        final JSONObject outerObject1 = new JSONObject(response1.getBody());
        final JSONArray assetsJsonArray1 = outerObject1.getJSONArray("Assets");
        double totalEstimate = 0;
        int itemListIndex = 0;
        while (itemListIndex < assetsJsonArray1.length()) {                                                                     // CREATE STORY/DEFECT DIRECTORIES + THEIR JSON FILES
            final JSONObject item = assetsJsonArray1.getJSONObject(itemListIndex);                                                            // read next story/defect in list response
            final JSONObject attributesObject= item.getJSONObject("Attributes");
            final JSONObject superObject= attributesObject.getJSONObject("Super");
            final JSONObject superValueObject = superObject.optJSONObject("value");
            if (!isProject || superValueObject == null){
                totalEstimate = totalEstimate + extractInformationOfStoryOrDefect(myExcelWorkbook, item);
            }
            itemListIndex++;
        }
        final ResponseEntity<String> response2 = getResponse(getURL(defectURL, id));                     // GET JSON RESPONSE LIST OF STORIES/DEFECTS IN EPIC
        final JSONObject outerObject2 = new JSONObject(response2.getBody());
        final JSONArray assetsJsonArray2 = outerObject2.getJSONArray("Assets");
        itemListIndex = 0;
        while (itemListIndex < assetsJsonArray2.length()) {                                                                     // CREATE STORY/DEFECT DIRECTORIES + THEIR JSON FILES
            final JSONObject item = assetsJsonArray2.getJSONObject(itemListIndex);                                                      // read next story/defect in list response
            final JSONObject attributesObject= item.getJSONObject("Attributes");
            final JSONObject superObject= attributesObject.getJSONObject("Super");
            final JSONObject superValueObject = superObject.optJSONObject("value");
            if (!isProject || superValueObject == null){
                totalEstimate = totalEstimate + extractInformationOfStoryOrDefect(myExcelWorkbook, item);
            }
            itemListIndex++;
        }
        return totalEstimate;
    }

    private static double extractInformationOfStoryOrDefect(final ExcelWorkbook myExcelWorkbook, final JSONObject item ) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final JSONObject itemAttributeObject = item.getJSONObject("Attributes");                                                       // get story/defect name
        double totalEstimate = 0;
        final JSONObject estimateObject = itemAttributeObject.getJSONObject("Estimate");
        final Double estimate = estimateObject.optDouble("value", -1);
        if (estimate != -1)
            totalEstimate = estimate;
        final String itemScopeToken = item.getString("href");                                                                     // get story/defect scope token
        final String itemID = item.getString("id");                                                                               //get story/defect id
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, itemScopeToken));                        // get story/defect json response
        final JSONObject outerObject = new JSONObject(response.getBody());
        final int rowNum = parseJSONObjectBodyToMapForExcel(myExcelWorkbook, outerObject, Sheet_Type.Story_Defect_Sheet);
        extractCreateData(myExcelWorkbook, itemScopeToken, Sheet_Type.Story_Defect_Sheet, rowNum);                       // GET DATA ON CREATED BY AND DATE
        extractClosedData(myExcelWorkbook, itemScopeToken, Sheet_Type.Story_Defect_Sheet, rowNum);                       // GET DATA ON CLOSED BY AND DATE
        extractDoneDate(myExcelWorkbook, itemScopeToken, Sheet_Type.Story_Defect_Sheet, rowNum);                         // GET DATA ON DONE DATE
        if (itemID.contains("Story"))
            extractAcceptanceCriteria(myExcelWorkbook, itemScopeToken, rowNum);                                          // GET ACCEPTANCE CRITERIA IF ITEM IS STORY
        extractTotalDoneHoursForStoryOrDefect(myExcelWorkbook, itemID, rowNum);
        extractInformationOfTasks(myExcelWorkbook, itemID, rowNum);                                                      // GET DATA ON TASKS IN STORY/DEFECT
        return totalEstimate;
    }

    private static void extractInformationOfTasks(final ExcelWorkbook myExcelWorkbook, final String itemID, final int itemRowNum) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Task_List_URL, itemID));                // GET JSON RESPONSE LIST OF TASKS IN STORY/DEFECT
        final JSONObject outerObjectTasks = new JSONObject(response.getBody());
        final JSONArray assetsJsonArrayTasks = outerObjectTasks.getJSONArray("Assets");
        extractTotalDetailEstimateForStoryOrDefect(myExcelWorkbook, assetsJsonArrayTasks, itemRowNum);                                       // GET DATA ON TOTAL DETAIL ESTIMATE FOR STORY/DEFECT
        extractTotalToDoHoursForStoryOrDefect(myExcelWorkbook, assetsJsonArrayTasks, itemRowNum);                                            // GET DATA ON TOTAL TO DO FOR STORY/DEFECT
        int taskListIndex = 0;
        while (taskListIndex < assetsJsonArrayTasks.length()) {                                                             // CREATE TASK DIRECTORIES + THEIR JSON FILES
            final JSONObject task = assetsJsonArrayTasks.getJSONObject(taskListIndex);                         // read next task in list response
            extractInformationOfTask(myExcelWorkbook, task);
            taskListIndex++;
        }
    }

    private static void extractInformationOfTask(final ExcelWorkbook myExcelWorkbook, final JSONObject task ) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String taskScopeToken = task.getString("href");                                                                    // get task scope token
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, taskScopeToken));                       // get task json response
        final JSONObject outerObject = new JSONObject(response.getBody());
        final int rowNum = parseJSONObjectBodyToMapForExcel(myExcelWorkbook, outerObject, Sheet_Type.Task_Sheet);
        extractCreateData(myExcelWorkbook, taskScopeToken, Sheet_Type.Task_Sheet, rowNum);                           // GET DATA ON CREATED BY AND DATE
        extractClosedData(myExcelWorkbook, taskScopeToken, Sheet_Type.Task_Sheet, rowNum);                           // GET DATA ON CLOSED BY AND DATE
        extractDoneHoursForTask(myExcelWorkbook, taskScopeToken, rowNum);                                            // GET DATA ON DONE HOURS
        extractInformationOfActuals(myExcelWorkbook, taskScopeToken, rowNum);                                        // GET DATA ON ACTUALS IN TASK
    }

    private static void extractInformationOfActuals(final ExcelWorkbook myExcelWorkbook, final String taskScopeToken, final int rowNum) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Actuals_List_URL, taskScopeToken)); // GET JSON RESPONSE LIST OF ACTUALS IN TASK
        final JSONObject outerObjectActuals = new JSONObject(response.getBody());
        final JSONObject attributesJsonObjectActuals = outerObjectActuals.getJSONObject("Attributes");
        final JSONObject actualsObject = attributesJsonObjectActuals.getJSONObject("Actuals");
        final JSONArray actualsArray = actualsObject.getJSONArray("value");

        int actualsListIndex = 0;
        while (actualsListIndex < actualsArray.length()) {                                                              // CREATE ACTUAL DIRECTORIES + THEIR JSON FILES
            final JSONObject actual = actualsArray.getJSONObject(actualsListIndex);                                               // read next actual in list response
            extractInformationOfActual(myExcelWorkbook, actual, actualsListIndex + 1, rowNum);
            actualsListIndex++;
        }
    }

    private static void extractInformationOfActual(final ExcelWorkbook myExcelWorkbook, final JSONObject actual, final int actualNumber, final int rowNum) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String actualID = actual.getString("idref");                                                       // get actual name
        final String actualScopeToken = actual.getString("href");                                                // get actual scope token
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, actualScopeToken));     // get actual json response
        final JSONObject outerObject = new JSONObject (response.getBody());
        final JSONObject attributesObject = outerObject.getJSONObject("Attributes");
        final JSONObject actualOwnerObject = attributesObject.getJSONObject("Member.Name");
        final String actualOwner = actualOwnerObject.getString("value");
        final JSONObject actualValueObject = attributesObject.getJSONObject("Value");
        final Double actualValue = actualValueObject.getDouble("value");
        final JSONObject actualDateObject = attributesObject.getJSONObject("Date");
        final String actualDate = actualDateObject.getString("value");
        Map<String, String> dataMap = new HashMap<>();
        dataMap.put("Actual " + actualNumber + " id" , actualID);
        dataMap.put("Actual " + actualNumber + " Owner" , actualOwner);
        dataMap.put("Actual " + actualNumber + " Value" , actualValue.toString());
        dataMap.put("Actual " + actualNumber + " Date" , actualDate);
        addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, Sheet_Type.Task_Sheet);
    }

    private static void extractEstimatePTS(final ExcelWorkbook myExcelWorkbook, final double total, final boolean isProject, final int rowNum, final Sheet_Type sheet_type ) throws Exception, IOException {
        final String name;
        if (isProject)
            name = "Estimate PTS - Rollup";
        else
            name = "Estimate PTS";
        Map<String, String> dataMap = new HashMap<>();
        dataMap.put(name, Double.toString(total));
        addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, sheet_type);
    }

    private static void extractTotalDoneHoursForStoryOrDefect(final ExcelWorkbook myExcelWorkbook, final String itemID, final int rowNum) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Done_Hours_List_URL, itemID));  // GET JSON RESPONSE DONE HOURS IN TASK
        final JSONObject outerObject = new JSONObject (response.getBody());
        final JSONArray assetsArray = outerObject.getJSONArray("Assets");

        int index = 0;
        double total = 0;
        while (index < assetsArray.length()) {
            final JSONObject taskTotalObject = assetsArray.getJSONObject(index);
            final JSONObject attributesObject = taskTotalObject.getJSONObject("Attributes");
            final JSONObject valueObject = attributesObject.getJSONObject("Actuals.Value.@Sum");
            final Double value = valueObject.optDouble("value", -1);
            if (value != -1)
                total = total + value;
            index++;
        }
        Map<String, String> dataMap = new HashMap<>();
        dataMap.put("Total Done Hours", Double.toString(total));
        addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, Sheet_Type.Story_Defect_Sheet);
    }

    private static void extractTotalToDoHoursForStoryOrDefect(final ExcelWorkbook myExcelWorkbook, final JSONArray assetsJsonArrayTasks, final int rowNum) throws Exception, IOException {
        int index = 0;
        double total = 0;
        while (index < assetsJsonArrayTasks.length()) {
            final JSONObject taskObject = assetsJsonArrayTasks.getJSONObject(index);
            final JSONObject attributesObject = taskObject.getJSONObject("Attributes");
            final JSONObject toDoObject = attributesObject.getJSONObject("ToDo");
            final Double toDo = toDoObject.optDouble("value", -1);
            if (toDo != -1)
                total = total + toDo;
            index++;
        }
        Map<String, String> dataMap = new HashMap<>();
        dataMap.put("Total To Do Hours", Double.toString(total));
        addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, Sheet_Type.Story_Defect_Sheet);
    }

    private static void extractTotalDetailEstimateForStoryOrDefect(final ExcelWorkbook myExcelWorkbook, final JSONArray assetsJsonArrayTasks, final int rowNum) throws Exception, IOException {
        int index = 0;
        double total = 0;
        while (index < assetsJsonArrayTasks.length()) {
            final JSONObject taskObject = assetsJsonArrayTasks.getJSONObject(index);
            final JSONObject attributesObject = taskObject.getJSONObject("Attributes");
            final JSONObject detailEstimateObject = attributesObject.getJSONObject("DetailEstimate");
            final Double detailEstimate = detailEstimateObject.optDouble("value", -1);
            if (detailEstimate != -1)
                total = total + detailEstimate;
            index++;
        }
        Map<String, String> dataMap = new HashMap<>();
        dataMap.put("Total Detail Estimate", Double.toString(total));
        addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, Sheet_Type.Story_Defect_Sheet);
    }

    private static void extractDoneHoursForTask(final ExcelWorkbook myExcelWorkbook, final String scopeToken, final int rowNum) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Done_Hours_URL, scopeToken));   // GET JSON RESPONSE DONE HOURS IN TASK
        final JSONObject outerObject = new JSONObject(response.getBody());
        final JSONObject attributeObject = outerObject.getJSONObject("Attributes");
        final JSONObject valueObject = attributeObject.getJSONObject("Actuals.Value.@Sum");
        final String name = valueObject.optString("value");
        if (name != null) {
            Map<String, String> dataMap = new HashMap<>();
            dataMap.put("Total Done Hours", name);
            addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, Sheet_Type.Task_Sheet);
        }
    }

    private static void extractCreateData(final ExcelWorkbook myExcelWorkbook, final String scopeToken, final Sheet_Type sheet_type, final int rowNum) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response1 = getResponse(getURL(Project_URL_Type.CreatedBy_URL, scopeToken));   // GET JSON RESPONSE CREATED BY
        final JSONObject outerObject = new JSONObject(response1.getBody());
        final JSONObject valueObject = outerObject.optJSONObject("value");
        if (valueObject != null) {
            final String memberScopeToken = valueObject.getString("href");
            final ResponseEntity<String> response2 = getResponse(getURL(Project_URL_Type.Base_URL, memberScopeToken));
            final JSONObject outerObjectMember = new JSONObject(response2.getBody());
            final JSONObject attributeObject = outerObjectMember.getJSONObject("Attributes");
            final JSONObject nameObject = attributeObject.getJSONObject("Name");
            final String name = nameObject.optString("value");
            if (name != null) {
                Map<String, String> dataMap = new HashMap<>();
                dataMap.put("CreatedBy", name);
                addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, sheet_type);
            }
        }
        final ResponseEntity<String> response3 = getResponse(getURL(Project_URL_Type.CreateDate_URL, scopeToken));  // GET JSON RESPONSE CREATE DATE
        final JSONObject outerObjectCreateDate = new JSONObject(response3.getBody());
        final String valueCreateDate = outerObjectCreateDate.optString("value");
        if (valueCreateDate != null) {
            Map<String, String> dataMap = new HashMap<>();
            dataMap.put("CreateDate", valueCreateDate);
            addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, sheet_type);
        }
    }

    private static void extractClosedData(final ExcelWorkbook myExcelWorkbook, final String scopeToken, final Sheet_Type sheet_type, final int rowNum) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response1 = getResponse(getURL(Project_URL_Type.ClosedBy_URL, scopeToken));                  // GET JSON RESPONSE CLOSED BY
        final JSONObject outerObject = new JSONObject(response1.getBody());
        final JSONObject valueObject = outerObject.optJSONObject("value");
        if (valueObject != null) {
            String memberScopeToken = valueObject.getString("href");
            final ResponseEntity<String> response2 = getResponse(getURL(Project_URL_Type.Base_URL, memberScopeToken));
            final JSONObject outerObjectMember = new JSONObject(response2.getBody());
            final JSONObject attributeObject = outerObjectMember.getJSONObject("Attributes");
            final JSONObject nameObject = attributeObject.getJSONObject("Name");
            final String name = nameObject.optString("value");
            if (name != null) {
                Map<String, String> dataMap = new HashMap<>();
                dataMap.put("ClosedBy", name);
                addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, sheet_type);
            }
            final ResponseEntity<String> response3 = getResponse(getURL(Project_URL_Type.ClosedDate_URL, scopeToken));            // GET JSON RESPONSE CLOSED DATE
            final JSONObject outerObjectClosedDate = new JSONObject(response3.getBody());
            final String valueClosedDate = outerObjectClosedDate.optString("value");
            if (valueClosedDate != null) {
                Map<String, String> dataMap = new HashMap<>();
                dataMap.put("ClosedDate", valueClosedDate);
                addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, sheet_type);
            }
        }
    }

    private static void extractDoneDate(final ExcelWorkbook myExcelWorkbook, final String scopeToken, final Sheet_Type sheet_type, final int rowNum) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.DoneDate_URL, scopeToken));   // GET JSON RESPONSE DONE DATE
        final JSONObject outerObject = new JSONObject(response.getBody());
        final String value = outerObject.optString("value");
        if (value != null) {
            Map<String, String> dataMap = new HashMap<>();
            dataMap.put("DoneDate", value);
            addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, sheet_type);
        }
    }

    private static void extractAcceptanceCriteria(final ExcelWorkbook myExcelWorkbook, final String scopeToken,final int rowNum) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.AcceptanceCriteria_URL, scopeToken));   // GET JSON RESPONSE DONE DATE
        final JSONObject outerObject = new JSONObject(response.getBody());
        final String value = outerObject.optString("value");
        if (value != null) {
            Map<String, String> dataMap = new HashMap<>();
            dataMap.put("AcceptanceCriteria", value);
            addDataToExistingRow(myExcelWorkbook, rowNum, dataMap, Sheet_Type.Story_Defect_Sheet);
        }
    }

    private static void runExtractionOnProjectData(final String projectIDNumber) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String targetDirectory = GLOBAL_DIRECTORY_PATH + "\\VersionOne Data Extract (Excel) - All Project " + projectIDNumber + " Data\\"; //todo make username adaptable
        createNextDirectory(targetDirectory);                    // CREATE PROJECT DIRECTORIES
        ExcelWorkbook myExcelWorkbook = new ExcelWorkbook();
        setUpExcelFile(myExcelWorkbook, true);
        System.out.println("Running...");
        String projectName = extractInformationOfParentProject(myExcelWorkbook, projectIDNumber);      // GET DATA ON PARENT PROJECT (URL IS HARD CODED)
        final FileOutputStream fileOut = new FileOutputStream(targetDirectory + makeAcceptableName(projectName) + " Extracted Data.xlsx");
        createExcelFile(myExcelWorkbook, fileOut);
    }

    private static void runExtractionOnChildProjectData(final String childProjectIDNumber) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String targetDirectory = GLOBAL_DIRECTORY_PATH + "\\VersionOne Data Extract (Excel) - Only Child Project " + childProjectIDNumber + " Data\\";
        createNextDirectory(targetDirectory);                                   // CREATE PROJECT DIRECTORIES
        ExcelWorkbook myExcelWorkbook = new ExcelWorkbook();
        setUpExcelFile(myExcelWorkbook, false);
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, "/KronosUltimate/rest-1.v1/Data/Scope/" + childProjectIDNumber));
        final JSONObject childProject = new JSONObject(response.getBody());
        final JSONObject attributeObject = childProject.getJSONObject("Attributes");
        final JSONObject nameObject = attributeObject.getJSONObject("Name");
        final String projectName = nameObject.optString("value");
        System.out.println("Running...");
        extractInformationOfChildProject(myExcelWorkbook, childProject, Sheet_Type.Parent_Project_Sheet);
        final FileOutputStream fileOut = new FileOutputStream(targetDirectory + makeAcceptableName(projectName) + " Extracted Data.xlsx");
        createExcelFile(myExcelWorkbook, fileOut);
    }

    private static int parseJSONObjectBodyToMapForExcel(final ExcelWorkbook myExcelWorkbook, final JSONObject outerObject, final Sheet_Type sheet_type) throws Exception {
        Map<String, String> dataMap = new HashMap<>();
        final String id = outerObject.getString("id");
        final JSONObject attributeObject = outerObject.getJSONObject("Attributes");
        for (String keyStr : attributeObject.keySet()) {
            JSONObject object = attributeObject.getJSONObject(keyStr);
            String value = handleObject(object);
            if (!value.isEmpty()) {
                dataMap.put(keyStr, value);
            }
        }
        dataMap.put("id", id);
        return addRowToSheet(myExcelWorkbook, dataMap, sheet_type);
    }

    private static String handleObject(final JSONObject object) {
        Object valueObject = object.opt("value");
        if (valueObject instanceof JSONArray || valueObject == null) {
            return "";
        } else if (valueObject instanceof JSONObject) {
            return ((JSONObject) valueObject).optString("idref");
        } else {
            return object.optString("value");
        }
    }

    private static int addRowToSheet (final ExcelWorkbook myExcelWorkbook, final Map<String, String> dataMap, final Sheet_Type sheet_type) throws Exception {
        final XSSFRow newRow;
        switch (sheet_type) {
            case Parent_Project_Sheet:
                newRow = myExcelWorkbook.wb.getSheet("Project").createRow(myExcelWorkbook.projectSheetRowIndex);
                myExcelWorkbook.projectSheetRowIndex++;
                break;
            case Child_Project_Sheet:
                newRow = myExcelWorkbook.wb.getSheet("Child Projects").createRow(myExcelWorkbook.childProjectSheetRowIndex);
                myExcelWorkbook.childProjectSheetRowIndex++;
                break;
            case Epic_Sheet:
                newRow = myExcelWorkbook.wb.getSheet("Epics").createRow(myExcelWorkbook.epicSheetRowIndex);
                myExcelWorkbook.epicSheetRowIndex++;
                break;
            case Story_Defect_Sheet:
                newRow = myExcelWorkbook.wb.getSheet("Stories and Defects").createRow(myExcelWorkbook.storyDefectSheetRowIndex);
                myExcelWorkbook.storyDefectSheetRowIndex++;
                break;
            case Task_Sheet:
                newRow = myExcelWorkbook.wb.getSheet("Tasks").createRow(myExcelWorkbook.taskSheetRowIndex);
                myExcelWorkbook.taskSheetRowIndex++;
                break;
            default:
                throw new Exception("Data sheet was not chosen correctly.");
        }
        addDataToNewRow(myExcelWorkbook, newRow, dataMap, sheet_type);
        return newRow.getRowNum();
    }

    private static void addDataToNewRow (final ExcelWorkbook myExcelWorkbook, final XSSFRow newRow, final Map<String, String> dataMap, final Sheet_Type sheet_type) throws Exception {
        for (Map.Entry<String, String> entry : dataMap.entrySet()) {
            XSSFCell newRowCell;
            switch (sheet_type) {
                case Parent_Project_Sheet:
                    if (!myExcelWorkbook.projectSheetColumns.contains(entry.getKey())) {
                        addColumnHeader(myExcelWorkbook, entry.getKey(), sheet_type);
                    }
                    newRowCell = newRow.createCell(myExcelWorkbook.projectSheetColumns.indexOf(entry.getKey()));
                    newRowCell.setCellValue(StringUtils.substring(entry.getValue(), 0, 32767));
                    break;
                case Child_Project_Sheet:
                    if (!myExcelWorkbook.childProjectSheetColumns.contains(entry.getKey())) {
                        addColumnHeader(myExcelWorkbook, entry.getKey(), sheet_type);
                    }
                    newRowCell = newRow.createCell(myExcelWorkbook.childProjectSheetColumns.indexOf(entry.getKey()));
                    newRowCell.setCellValue(StringUtils.substring(entry.getValue(), 0, 32767));
                    break;
                case Epic_Sheet:
                    if (!myExcelWorkbook.epicSheetColumns.contains(entry.getKey())) {
                        addColumnHeader(myExcelWorkbook, entry.getKey(), sheet_type);
                    }
                    newRowCell = newRow.createCell(myExcelWorkbook.epicSheetColumns.indexOf(entry.getKey()));
                    newRowCell.setCellValue(StringUtils.substring(entry.getValue(), 0, 32767));
                    break;
                case Story_Defect_Sheet:
                    if (!myExcelWorkbook.storyDefectSheetColumns.contains(entry.getKey())) {
                        addColumnHeader(myExcelWorkbook, entry.getKey(), sheet_type);
                    }
                    newRowCell = newRow.createCell(myExcelWorkbook.storyDefectSheetColumns.indexOf(entry.getKey()));
                    newRowCell.setCellValue(StringUtils.substring(entry.getValue(), 0, 32767));
                    break;
                case Task_Sheet:
                    if (!myExcelWorkbook.taskSheetColumns.contains(entry.getKey())) {
                        addColumnHeader(myExcelWorkbook, entry.getKey(), sheet_type);
                    }
                    newRowCell = newRow.createCell(myExcelWorkbook.taskSheetColumns.indexOf(entry.getKey()));
                    newRowCell.setCellValue(StringUtils.substring(entry.getValue(), 0, 32767));
                    break;
                default:
                    throw new Exception("Data sheet was not chosen correctly.");
            }
        }
    }

    private static void addDataToExistingRow (final ExcelWorkbook myExcelWorkbook, final int rowNum, final Map<String, String> dataMap, final Sheet_Type sheet_type) throws Exception {
        for (Map.Entry<String, String> entry : dataMap.entrySet()) {
            XSSFCell newRowCell;
            switch (sheet_type) {
                case Parent_Project_Sheet:
                    if (!myExcelWorkbook.projectSheetColumns.contains(entry.getKey())) {
                        addColumnHeader(myExcelWorkbook, entry.getKey(), sheet_type);
                    }
                    newRowCell = myExcelWorkbook.wb.getSheet("Project").getRow(rowNum).createCell(myExcelWorkbook.projectSheetColumns.indexOf(entry.getKey()));
                    newRowCell.setCellValue(StringUtils.substring(entry.getValue(), 0, 32767));
                    break;
                case Child_Project_Sheet:
                    if (!myExcelWorkbook.childProjectSheetColumns.contains(entry.getKey())) {
                        addColumnHeader(myExcelWorkbook, entry.getKey(), sheet_type);
                    }
                    newRowCell = myExcelWorkbook.wb.getSheet("Child Projects").getRow(rowNum).createCell(myExcelWorkbook.childProjectSheetColumns.indexOf(entry.getKey()));
                    newRowCell.setCellValue(StringUtils.substring(entry.getValue(), 0, 32767));
                    break;
                case Epic_Sheet:
                    if (!myExcelWorkbook.epicSheetColumns.contains(entry.getKey())) {
                        addColumnHeader(myExcelWorkbook, entry.getKey(), sheet_type);
                    }
                    newRowCell = myExcelWorkbook.wb.getSheet("Epics").getRow(rowNum).createCell(myExcelWorkbook.epicSheetColumns.indexOf(entry.getKey()));
                    newRowCell.setCellValue(StringUtils.substring(entry.getValue(), 0, 32767));
                    break;
                case Story_Defect_Sheet:
                    if (!myExcelWorkbook.storyDefectSheetColumns.contains(entry.getKey())) {
                        addColumnHeader(myExcelWorkbook, entry.getKey(), sheet_type);
                    }
                    newRowCell =  myExcelWorkbook.wb.getSheet("Stories and Defects").getRow(rowNum).createCell(myExcelWorkbook.storyDefectSheetColumns.indexOf(entry.getKey()));
                    newRowCell.setCellValue(StringUtils.substring(entry.getValue(), 0, 32767));
                    break;
                case Task_Sheet:
                    if (!myExcelWorkbook.taskSheetColumns.contains(entry.getKey())) {
                        addColumnHeader(myExcelWorkbook, entry.getKey(), sheet_type);
                    }
                    newRowCell = myExcelWorkbook.wb.getSheet("Tasks").getRow(rowNum).createCell(myExcelWorkbook.taskSheetColumns.indexOf(entry.getKey()));
                    newRowCell.setCellValue(StringUtils.substring(entry.getValue(), 0, 32767));
                    break;
                default:
                    throw new Exception("Data sheet was not chosen correctly.");
            }
        }
    }

    private static void addColumnHeader(final ExcelWorkbook myExcelWorkbook, final String key, final Sheet_Type sheet_type) throws Exception {
        final XSSFFont columnFont = myExcelWorkbook.wb.createFont();
        columnFont.setBold(true);
        final XSSFCellStyle columnStyle = myExcelWorkbook.wb.createCellStyle();
        columnStyle.setFont(columnFont);
        XSSFCell newRowCell;
        switch(sheet_type) {
            case Parent_Project_Sheet:
                myExcelWorkbook.projectSheetColumns.add(key);
                newRowCell = myExcelWorkbook.wb.getSheet("Project").getRow(0).createCell(myExcelWorkbook.projectSheetColumns.indexOf(key));
                break;
            case Child_Project_Sheet:
                myExcelWorkbook.childProjectSheetColumns.add(key);
                newRowCell = myExcelWorkbook.wb.getSheet("Child Projects").getRow(0).createCell(myExcelWorkbook.childProjectSheetColumns.indexOf(key));
                break;
            case Epic_Sheet:
                myExcelWorkbook.epicSheetColumns.add(key);
                newRowCell = myExcelWorkbook.wb.getSheet("Epics").getRow(0).createCell(myExcelWorkbook.epicSheetColumns.indexOf(key));
                break;
            case Story_Defect_Sheet:
                myExcelWorkbook.storyDefectSheetColumns.add(key);
                newRowCell = myExcelWorkbook.wb.getSheet("Stories and Defects").getRow(0).createCell(myExcelWorkbook.storyDefectSheetColumns.indexOf(key));
                break;
            case Task_Sheet:
                myExcelWorkbook.taskSheetColumns.add(key);
                newRowCell = myExcelWorkbook.wb.getSheet("Tasks").getRow(0).createCell(myExcelWorkbook.taskSheetColumns.indexOf(key));
                break;
            default:
                throw new Exception("Data sheet was not chosen correctly.");
        }
        newRowCell.setCellStyle(columnStyle);
        newRowCell.setCellValue(key);
    }

    private static void setUpExcelFile (final ExcelWorkbook myExcelWorkbook, final boolean hasChildProjects) {
        XSSFSheet projectSheet = myExcelWorkbook.wb.createSheet("Project");   // CREATE SHEET IN WB
        projectSheet.createRow(0);                                              // CREATE ROW FOR HEADERS IN EACH SHEET
        if (hasChildProjects) {
            XSSFSheet childProjectSheet = myExcelWorkbook.wb.createSheet("Child Projects");
            childProjectSheet.createRow(0);
        }
        XSSFSheet epicSheet = myExcelWorkbook.wb.createSheet("Epics");
        epicSheet.createRow(0);
        XSSFSheet storyDefectSheet = myExcelWorkbook.wb.createSheet("Stories and Defects");
        storyDefectSheet.createRow(0);
        XSSFSheet taskSheet = myExcelWorkbook.wb.createSheet("Tasks");
        taskSheet.createRow(0);
    }

    private static void createExcelFile (final ExcelWorkbook myExcelWorkbook, final FileOutputStream fileOut) throws IOException {
        myExcelWorkbook.wb.write(fileOut);
        myExcelWorkbook.wb.close();
        fileOut.flush();
        fileOut.close();
        System.out.println("Excel file created successfully.");
    }

    public static void main(String[] args) {
        try {
            // ENTER BEARER AUTHENTICATION OBTAINED FROM VERSIONONE
                setBearerAuthToken("exampleBearerAuthentication");

            // ENTER COMPANY/TEAM/DOMAIN NAME IN VERSIONONE URL
                setCompany("exampleCompany");

            // CASES - must provide item ID number if you want a specific item
                runExtractionOnProjectData("example");
                runExtractionOnChildProjectData("example");

        } catch (
            /*KeyManagementException | NoSuchAlgorithmException | KeyStoreException | IOException |*/ Exception e) {
            e.printStackTrace();
        }
    }
}
