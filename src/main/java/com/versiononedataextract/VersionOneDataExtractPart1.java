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
import org.apache.commons.lang3.StringUtils;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.http.*;
import org.springframework.http.client.HttpComponentsClientHttpRequestFactory;
import org.springframework.web.client.RestTemplate;

import javax.net.ssl.SSLContext;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.KeyManagementException;
import java.security.KeyStoreException;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.List;

public class VersionOneDataExtractPart1 {

    private static final String GLOBAL_DIRECTORY_PATH = System.getProperty("user.home") + File.separatorChar + "Documents";

    private static String bearerAuthToken;

    private static String company;

    private static void setBearerAuthToken(final String token) {
        bearerAuthToken = token;
    }

    private static void setCompany(final String companyToken) {
        company = companyToken;
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
        switch (project_URL_Type){
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

    private static void createFile(final String target, final String content) throws IOException {
        final File file = new File(target);
        final FileWriter fileWriter = new FileWriter(file);
        final PrintWriter printWriter = new PrintWriter(fileWriter);
        printWriter.print(content);
        fileWriter.flush();
        fileWriter.close();
        System.out.println("File created successfully: " + target);
    }

    private static String makeAcceptableName(final String name) {
        return StringUtils.left(name.replaceAll("[^a-zA-Z0-9\\-]", "_"), 127);
    }

    private static void extractInformationOfParentProject(final String projectIDNumber, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.My_Project_URL, projectIDNumber));  // GET PROJECT JSON RESPONSE (URL IS HARD CODED)
        final JSONObject outerObject = new JSONObject(response.getBody());
        final String projectScopeToken = outerObject.getString("href");
        final JSONObject projectAttributeObject = outerObject.getJSONObject("Attributes");                    // get child name
        final JSONObject projectNameObject = projectAttributeObject.getJSONObject("Name");
        final String projectName = projectNameObject.getString("value");
        final String parentTargetDirectory = targetDirectory + projectName + "\\";
        createNextDirectory(parentTargetDirectory);
        createFile(parentTargetDirectory + projectName + ".json", response.getBody());                           // OUTPUT THE PROJECT JSON RESPONSE TO A JSON FILE
        extractCreateData(projectScopeToken, parentTargetDirectory);                                                    // GET DATA ON CREATED BY AND DATE
        extractClosedData(projectScopeToken, parentTargetDirectory);                                                    // GET DATA ON CLOSED BY AND DATE
        double totalEstimate = extractInformationOfStoriesAndDefects(true, "Scope:" + projectIDNumber, parentTargetDirectory);
        totalEstimate = totalEstimate + extractInformationOfChildProjects(projectIDNumber, parentTargetDirectory);      // GET DATA ON CHILD PROJECTS IN PROJECT (URL IS HARD CODED)
        extractEstimatePTS(totalEstimate, parentTargetDirectory, true);                              // create done hours json file in the directory
    }

    private static double extractInformationOfChildProjects(final String projectIDNumber, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.My_Child_Project_List_URL, projectIDNumber));    // GET JSON RESPONSE LIST OF CHILD PROJECTS (URL IS HARD CODED)
        final JSONObject outerObjectChildren = new JSONObject(response.getBody());
        final JSONArray assetsJsonArrayChildren = outerObjectChildren.getJSONArray("Assets");

        double totalEstimate = 0;
        int childProjectIndex = 0;
        while (childProjectIndex < assetsJsonArrayChildren.length()) {                                                              // CREATE CHILD PROJECT DIRECTORIES + THEIR JSON FILES
            final JSONObject childProject = assetsJsonArrayChildren.getJSONObject(childProjectIndex);                       // read next child project in list
            totalEstimate = totalEstimate + extractInformationOfChildProject(childProject, targetDirectory);
            childProjectIndex++;
        }
        return totalEstimate;
    }

    private static double extractInformationOfChildProject(final JSONObject childProject, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final JSONObject childProjectAttributeObject = childProject.getJSONObject("Attributes");                                // get child name
        final JSONObject childProjectNameObject = childProjectAttributeObject.getJSONObject("Name");
        final String childProjectName = childProjectNameObject.getString("value");
        final String childProjectScopeToken = childProject.getString("href");                                               // get child scope token
        final String childProjectID = childProject.getString("id");                                                         // get child id
        final String childTargetDirectory = targetDirectory + makeAcceptableName(childProjectName) + "\\";                      // use original target directory
        createNextDirectory(childTargetDirectory);                                                                              // create child project directory
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, childProjectScopeToken));         // get child project json response
        createFile(childTargetDirectory + makeAcceptableName(childProjectName) + ".json", response.getBody());           // create child project json file in the child project directory
        extractCreateData(childProjectScopeToken, childTargetDirectory);                                             // GET DATA ON CREATED BY AND DATE
        extractClosedData(childProjectScopeToken, childTargetDirectory);                                             // GET DATA ON CLOSED BY AND DATE
        double totalEstimate = extractInformationOfStoriesAndDefects(true,childProjectID, childTargetDirectory);
        totalEstimate = totalEstimate + extractInformationOfEpics(childProjectID, childTargetDirectory);             // GET DATA ON EPICS IN CHILD PROJECT
        extractEstimatePTS(totalEstimate, childTargetDirectory, true);
        return totalEstimate;
    }

    private static double extractInformationOfEpics(final String childProjectID, final String childTargetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
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
                totalEstimate = totalEstimate + extractInformationOfEpic(epic, childProjectID, childTargetDirectory);
            }
            epicIndex++;
        }
        return totalEstimate;
    }

    private static double extractInformationOfEpic(final JSONObject epic, final String childProjectIDForAssertion, final String childTargetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        double totalEstimate = 0;
        final JSONObject epicAttributeObject = epic.getJSONObject("Attributes");
        final JSONObject epicNameObject = epicAttributeObject.getJSONObject("Name");                                                            // get epic name
        final String epicName = epicNameObject.getString("value");
        final JSONObject epicNumberObject = epicAttributeObject.getJSONObject("Number");                                                        // get epic number
        final String epicNumber = epicNumberObject.getString("value");
        final String epicScopeToken = epic.getString("href");                                                                              // get epic scope token
        final String epicID = epic.getString("id");                                                                                        //get epic id
        final String epicTargetDirectory = childTargetDirectory + epicNumber + " " + makeAcceptableName(epicName) + "\\";                       // use child project directory
        createNextDirectory(epicTargetDirectory);                                                                                               // create epic directory
        final ResponseEntity<String> response1 = getResponse(getURL(Project_URL_Type.Base_URL, epicScopeToken));                                // get epic json response
        createFile(epicTargetDirectory + epicNumber + " " + makeAcceptableName(epicName) + ".json", response1.getBody());                // create epic json file in the epic directory
        extractCreateData(epicScopeToken, epicTargetDirectory);                                                                 // GET DATA ON CREATED BY AND DATE
        extractClosedData(epicScopeToken, epicTargetDirectory);                                                                 // GET DATA ON CLOSED BY AND DATE
        extractDoneDate(epicScopeToken, epicTargetDirectory);                                                                   // GET DATA ON DONE DATE
        String childProjectID = childProjectIDForAssertion;
        if(childProjectIDForAssertion.equals("notAnID")) {
            final JSONObject epicSecurityScopeObject = epicAttributeObject.getJSONObject("SecurityScope");                                      // get epic security scope
            final JSONObject epicSecurityScopeValueObject = epicSecurityScopeObject.getJSONObject("value");
            childProjectID = epicSecurityScopeValueObject.getString("idref");
        }
        totalEstimate = totalEstimate + extractInformationOfStoriesAndDefects(false, epicID, epicTargetDirectory);      // GET DATA ON DEFECTS AND STORIES IN SUB-EPIC
        final ResponseEntity<String> response2 = getResponse(getURL(Project_URL_Type.Epic_List_URL, childProjectID));           // GET JSON RESPONSE LIST OF EPICS IN CHILD PROJECT
        final JSONObject outerObjectEpics = new JSONObject(response2.getBody());
        final JSONArray assetsJsonArrayEpics = outerObjectEpics.getJSONArray("Assets");
        totalEstimate = totalEstimate + extractInformationOfSubEpics(assetsJsonArrayEpics, childProjectID, epicID, epicTargetDirectory);   // GET DATA ON SUB-EPICS IN EPIC
        extractEstimatePTS(totalEstimate, epicTargetDirectory, false);
        return totalEstimate;
    }

    private static double extractInformationOfSubEpics(final JSONArray assetsJsonArrayEpics, final String childProjectID, final String epicID, final String epicTargetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
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
                    final JSONObject subEpicNameObject = subEpicAttributeObject.getJSONObject("Name");                                                       // get sub-epic name
                    final String subEpicName = subEpicNameObject.getString("value");
                    final JSONObject subEpicNumberObject = subEpicAttributeObject.getJSONObject("Number");                                                   // get sub-epic number
                    final String subEpicNumber = subEpicNumberObject.getString("value");
                    final String subEpicScopeToken = subEpic.getString("href");                                                                         // get sub-epic scope token
                    final String subEpicID = subEpic.getString("id");                                                                                   //get sub-epic id`
                    final String subEpicTargetDirectory = epicTargetDirectory + subEpicNumber + " " + makeAcceptableName(subEpicName) + "\\";                // use parent epic directory
                    createNextDirectory(subEpicTargetDirectory);                                                                                             // create sub-epic directory
                    final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, subEpicScopeToken));                               // get sub-epic json response
                    createFile(subEpicTargetDirectory + subEpicNumber + " " + makeAcceptableName(subEpicName) + ".json", response.getBody());         // create sub-epic json file in the epic directory
                    extractCreateData(subEpicScopeToken, subEpicTargetDirectory);                                                           // GET DATA ON CREATED BY AND DATE
                    extractClosedData(subEpicScopeToken, subEpicTargetDirectory);                                                           // GET DATA ON CLOSED BY AND DATE
                    extractDoneDate(subEpicScopeToken, subEpicTargetDirectory);                                                             // GET DATA ON DONE DATE
                    double subTotalEstimate = extractInformationOfStoriesAndDefects(false, subEpicID, subEpicTargetDirectory) + extractInformationOfSubEpics(assetsJsonArrayEpics, childProjectID, subEpicID, subEpicTargetDirectory);;
                    extractEstimatePTS(subTotalEstimate, subEpicTargetDirectory, false);
                    totalEstimate = totalEstimate + subTotalEstimate;                                                                       // GET DATA ON DEFECTS AND STORIES IN SUB-EPIC AND DATA ON SUB-EPICS OF CURRENT CHILD EPIC
                }
            }
            subEpicIndex++;
        }
        return totalEstimate;
    }

    private static double extractInformationOfStoriesAndDefects(final boolean isProject, final String id, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final Project_URL_Type storyURL;
        final Project_URL_Type defectURL;
        if (isProject) {
            storyURL = Project_URL_Type.Story_List_In_Scope;
            defectURL = Project_URL_Type.Defect_List_In_Scope;
        } else {
            storyURL = Project_URL_Type.Story_List_URL;
            defectURL = Project_URL_Type.Defect_List_URL;
        }
        final ResponseEntity<String> response1 = getResponse(getURL(storyURL, id));                   // GET JSON RESPONSE LIST OF STORIES/DEFECTS IN EPIC
        final JSONObject outerObject1 = new JSONObject(response1.getBody());
        final JSONArray assetsJsonArray1 = outerObject1.getJSONArray("Assets");

        double totalEstimate = 0;
        int itemListIndex = 0;
        while (itemListIndex < assetsJsonArray1.length()) {                                          // CREATE STORY/DEFECT DIRECTORIES + THEIR JSON FILES
            final JSONObject item = assetsJsonArray1.getJSONObject(itemListIndex);                                // read next story/defect in list response
            final JSONObject attributesObject= item.getJSONObject("Attributes");
            final JSONObject superObject= attributesObject.getJSONObject("Super");
            final JSONObject superValueObject = superObject.optJSONObject("value");
            if (!isProject || superValueObject == null){
                totalEstimate = totalEstimate + extractInformationOfStoryOrDefect(item, targetDirectory);
            }
            itemListIndex++;
        }
        final ResponseEntity<String> response2 = getResponse(getURL(defectURL, id));                // GET JSON RESPONSE LIST OF STORIES/DEFECTS IN EPIC
        final JSONObject outerObject2 = new JSONObject(response2.getBody());
        final JSONArray assetsJsonArray2 = outerObject2.getJSONArray("Assets");

        itemListIndex = 0;
        while (itemListIndex < assetsJsonArray2.length()) {                                         // CREATE STORY/DEFECT DIRECTORIES + THEIR JSON FILES
            final JSONObject item = assetsJsonArray2.getJSONObject(itemListIndex);                         // read next story/defect in list response
            final JSONObject attributesObject= item.getJSONObject("Attributes");
            final JSONObject superObject= attributesObject.getJSONObject("Super");
            final JSONObject superValueObject = superObject.optJSONObject("value");
            if (!isProject || superValueObject == null){
                totalEstimate = totalEstimate + extractInformationOfStoryOrDefect(item, targetDirectory);
            }
            itemListIndex++;
        }
        return totalEstimate;
    }

    private static double extractInformationOfStoryOrDefect(final JSONObject item, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final JSONObject itemAttributeObject = item.getJSONObject("Attributes");                                                         // get story/defect name
        double totalEstimate = 0;
        final JSONObject estimateObject = itemAttributeObject.getJSONObject("Estimate");
        final Double estimate = estimateObject.optDouble("value", -1);
        if (estimate != -1)
            totalEstimate = estimate;
        final JSONObject itemNameObject = itemAttributeObject.getJSONObject("Name");
        final String itemName = itemNameObject.getString("value");
        final JSONObject itemNumberObject = itemAttributeObject.getJSONObject("Number");                                                 // get story/defect number
        final String itemNumber = itemNumberObject.getString("value");
        final String itemScopeToken = item.getString("href");                                                                       // get story/defect scope token
        final String itemID = item.getString("id");                                                                                 //get story/defect id
        final String itemTargetDirectory = targetDirectory + itemNumber + " " + makeAcceptableName(itemName) + "\\";                      // use child target directory
        createNextDirectory(itemTargetDirectory);                                                                                        // create story/defect directory
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, itemScopeToken));                          // get story/defect json response
        createFile(itemTargetDirectory + itemNumber + " " + makeAcceptableName(itemName) + ".json", response.getBody());          // create story json file in the story directory
        extractCreateData(itemScopeToken, itemTargetDirectory);                                                             // GET DATA ON CREATED BY AND DATE
        extractClosedData(itemScopeToken, itemTargetDirectory);                                                             // GET DATA ON CLOSED BY AND DATE
        extractDoneDate(itemScopeToken, itemTargetDirectory);                                                               // GET DATA ON DONE DATE
        if (itemID.contains("Story"))
            extractAcceptanceCriteria(itemScopeToken, itemTargetDirectory);                                                 // GET ACCEPTANCE CRITERIA IF ITEM IS STORY
        extractTotalDoneHoursForStoryOrDefect(itemID, itemTargetDirectory);
        extractInformationOfTasks(itemID, itemTargetDirectory);                                                             // GET DATA ON TASKS IN STORY/DEFECT
        return totalEstimate;
    }

    private static void extractInformationOfTasks(final String itemID, final String itemTargetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Task_List_URL, itemID));                 // GET JSON RESPONSE LIST OF TASKS IN STORY/DEFECT
        final JSONObject outerObjectTasks = new JSONObject(response.getBody());
        final JSONArray assetsJsonArrayTasks = outerObjectTasks.getJSONArray("Assets");
        extractTotalDetailEstimateForStoryOrDefect(assetsJsonArrayTasks, itemTargetDirectory);                              // GET DATA ON TOTAL DETAIL ESTIMATE FOR STORY/DEFECT
        extractTotalToDoHoursForStoryOrDefect(assetsJsonArrayTasks, itemTargetDirectory);                                   // GET DATA ON TOTAL TO DO FOR STORY/DEFECT

        int taskListIndex = 0;
        while (taskListIndex < assetsJsonArrayTasks.length()) {                                                             // CREATE TASK DIRECTORIES + THEIR JSON FILES
            final JSONObject task = assetsJsonArrayTasks.getJSONObject(taskListIndex);                                                 // read next task in list response
            extractInformationOfTask(task, itemTargetDirectory);
            taskListIndex++;
        }
    }

    private static void extractInformationOfTask(final JSONObject task, final String itemTargetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final JSONObject taskAttributeObject = task.getJSONObject("Attributes");                                                      // get task name
        final JSONObject taskNameObject = taskAttributeObject.getJSONObject("Name");
        final String taskName = taskNameObject.getString("value");
        final JSONObject taskNumberObject = taskAttributeObject.getJSONObject("Number");                                              // get task number
        final String taskNumber = taskNumberObject.getString("value");
        final String taskScopeToken = task.getString("href");                                                                    // get task scope token
        final String taskTargetDirectory = itemTargetDirectory + taskNumber + " " + makeAcceptableName(taskName) + "\\";              // use story/defect target directory
        createNextDirectory(taskTargetDirectory);                                                                                     // create task directory
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, taskScopeToken));                       // get task json response
        createFile(taskTargetDirectory + taskNumber + " " + makeAcceptableName(taskName) + ".json", response.getBody());       // create task json file in the task directory
        extractCreateData(taskScopeToken, taskTargetDirectory);                                                         // GET DATA ON CREATED BY AND DATE
        extractClosedData(taskScopeToken, taskTargetDirectory);                                                         // GET DATA ON CLOSED BY AND DATE
        extractDoneHoursForTask(taskScopeToken, taskTargetDirectory);                                                   // GET DATA ON DONE HOURS
        extractInformationOfActuals(taskScopeToken, taskTargetDirectory);                                               // GET DATA ON ACTUALS IN TASK
    }

    private static void extractInformationOfActuals(final String taskScopeToken, final String taskTargetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Actuals_List_URL, taskScopeToken)); // GET JSON RESPONSE LIST OF ACTUALS IN TASK
        final JSONObject outerObjectActuals = new JSONObject(response.getBody());
        final JSONObject attributesJsonObjectActuals = outerObjectActuals.getJSONObject("Attributes");
        final JSONObject actualsObject = attributesJsonObjectActuals.getJSONObject("Actuals");
        final JSONArray actualsArray = actualsObject.getJSONArray("value");

        int actualsListIndex = 0;
        while (actualsListIndex < actualsArray.length()) {                                                              // CREATE ACTUAL DIRECTORIES + THEIR JSON FILES
            final JSONObject actual = actualsArray.getJSONObject(actualsListIndex);                                               // read next actual in list response
            extractInformationOfActual(actual, taskTargetDirectory);
            actualsListIndex++;
        }
    }

    private static void extractInformationOfActual(final JSONObject actual, final String taskTargetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String actualName = actual.getString("idref");                                                     // get actual name
        final String actualScopeToken = actual.getString("href");                                                // get actual scope token
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, actualScopeToken));     // get actual json response
        createFile(taskTargetDirectory + makeAcceptableName(actualName) + ".json", response.getBody());        // create actual json file in the task directory
    }

    private static void extractEstimatePTS(final double total, final String targetDirectory, final boolean isProject) throws  IOException {
        final String fileName;
        if (isProject)
            fileName = "Estimate PTS - Rollup.json";
        else
            fileName = "Estimate PTS.json";
        createFile(targetDirectory + fileName, "{\"estimatePTS\": " + total + "}");
    }

    private static void extractTotalDoneHoursForStoryOrDefect(final String itemID, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
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
        createFile(targetDirectory + "Total Done Hours.json", "{\"totalDoneHours\": " + total + "}");       // create done hours json file in the directory
    }

    private static void extractTotalToDoHoursForStoryOrDefect(final JSONArray assetsJsonArrayTasks, final String targetDirectory) throws IOException {
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
        createFile(targetDirectory + "Total To Do Hours.json", "{\"totalToDo\": " + total + "}");
    }

    private static void extractTotalDetailEstimateForStoryOrDefect(final JSONArray assetsJsonArrayTasks, final String targetDirectory) throws IOException {
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
        createFile(targetDirectory + "Total Detail Estimate.json", "{\"totalDetailEstimate\": " + total + "}");
    }

    private static void extractDoneHoursForTask(final String scopeToken, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Done_Hours_URL, scopeToken));   // GET JSON RESPONSE DONE HOURS IN TASK
        createFile(targetDirectory + "Total Done Hours.json", response.getBody());                                   // create done hours json file in the directory
    }

    private static void extractCreateData(final String scopeToken, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response1 = getResponse(getURL(Project_URL_Type.CreatedBy_URL, scopeToken));   // GET JSON RESPONSE CREATED BY
        final JSONObject outerObject = new JSONObject(response1.getBody());
        final JSONObject valueObject = outerObject.optJSONObject("value");
        if (valueObject != null) {
            final String memberScopeToken = valueObject.getString("href");
            final ResponseEntity<String> response2 = getResponse(getURL(Project_URL_Type.Base_URL, memberScopeToken));
            createFile(targetDirectory + "CreatedBy.json", response2.getBody());                                   // created by json file in the directory
        }
        final ResponseEntity<String> response3 = getResponse(getURL(Project_URL_Type.CreateDate_URL, scopeToken));  // GET JSON RESPONSE CREATE DATE
        createFile(targetDirectory + "CreateDate.json", response3.getBody());                                      // create date json file in the directory
    }

    private static void extractClosedData(final String scopeToken, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response1 = getResponse(getURL(Project_URL_Type.ClosedBy_URL, scopeToken));                  // GET JSON RESPONSE CLOSED BY
        final JSONObject outerObject = new JSONObject(response1.getBody());
        final JSONObject valueObject = outerObject.optJSONObject("value");
        if (valueObject != null) {
            String memberScopeToken = valueObject.getString("href");
            final ResponseEntity<String> response2 = getResponse(getURL(Project_URL_Type.Base_URL, memberScopeToken));
            createFile(targetDirectory + "ClosedBy.json", response2.getBody());                                                  // closed by json file in the directory
            final ResponseEntity<String> response3 = getResponse(getURL(Project_URL_Type.ClosedDate_URL, scopeToken));            // GET JSON RESPONSE CLOSED DATE
            createFile(targetDirectory + "ClosedDate.json", response3.getBody());                                                // closed date hours json file in the directory
        }
    }

    private static void extractDoneDate(final String scopeToken, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.DoneDate_URL, scopeToken));   // GET JSON RESPONSE DONE DATE
        final JSONObject outerObject = new JSONObject(response.getBody());
        final JSONObject valueObject = outerObject.optJSONObject("value");
        if (valueObject != null)
            createFile(targetDirectory + "DoneDate.json", response.getBody());                                  // done date json file in the directory
    }

    private static void extractAcceptanceCriteria(final String scopeToken, final String targetDirectory) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.AcceptanceCriteria_URL, scopeToken));   // GET JSON RESPONSE DONE DATE
        createFile(targetDirectory + "AcceptanceCriteria.json", response.getBody());                                     // done date json file in the directory
    }

    private static void runExtractionOnProjectData(final String projectIDNumber) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String targetDirectory = GLOBAL_DIRECTORY_PATH + "\\VersionOne Data Extract - All Project " + projectIDNumber + " Data\\"; //todo make username adaptable
        createNextDirectory(targetDirectory);                                    // CREATE PROJECT DIRECTORIES
        extractInformationOfParentProject(projectIDNumber, targetDirectory);     // GET DATA ON PARENT PROJECT (URL IS HARD CODED)
    }

    private static void runExtractionOnChildProjectData(final String childProjectIDNumber) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String targetDirectory = GLOBAL_DIRECTORY_PATH + "\\VersionOne Data Extract - Only Child Project " + childProjectIDNumber + " Data\\";
        createNextDirectory(targetDirectory);                   // CREATE PROJECT DIRECTORIES
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, "/KronosUltimate/rest-1.v1/Data/Scope/" + childProjectIDNumber));
        final JSONObject childProject = new JSONObject(response.getBody());
        extractInformationOfChildProject(childProject, targetDirectory);
    }

    private static void runExtractionOnEpicData(final String epicIDNumber) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String targetDirectory = GLOBAL_DIRECTORY_PATH + "\\VersionOne Data Extract - Only Epic " + epicIDNumber + " Data\\";
        createNextDirectory(targetDirectory);                   // CREATE PROJECT DIRECTORIES
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, "/KronosUltimate/rest-1.v1/Data/Epic/" + epicIDNumber));
        final JSONObject epic = new JSONObject(response.getBody());
        extractInformationOfEpic(epic, "notAnID", targetDirectory);
    }

    private static void runExtractionOnStoryData(final String storyIDNumber) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String targetDirectory = GLOBAL_DIRECTORY_PATH + "\\VersionOne Data Extract - Only Story " + storyIDNumber + " Data\\";
        createNextDirectory(targetDirectory);                   // CREATE PROJECT DIRECTORIES
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, "/KronosUltimate/rest-1.v1/Data/Story/" + storyIDNumber));
        final JSONObject item = new JSONObject(response.getBody());
        extractInformationOfStoryOrDefect(item, targetDirectory);
    }

    private static void runExtractionOnDefectData(final String defectIDNumber) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String targetDirectory = GLOBAL_DIRECTORY_PATH + "\\VersionOne Data Extract - Only Defect " + defectIDNumber + " Data\\";
        createNextDirectory(targetDirectory);                   // CREATE PROJECT DIRECTORIES
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, "/KronosUltimate/rest-1.v1/Data/Defect/" + defectIDNumber));
        final JSONObject item = new JSONObject(response.getBody());
        extractInformationOfStoryOrDefect(item, targetDirectory);
    }

    private static void runExtractionOnTaskData(final String taskIDNumber) throws Exception, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, IOException {
        final String targetDirectory = GLOBAL_DIRECTORY_PATH + "\\VersionOne Data Extract - Only Task " + taskIDNumber + " Data\\";
        createNextDirectory(targetDirectory);                   // CREATE PROJECT DIRECTORIES
        final ResponseEntity<String> response = getResponse(getURL(Project_URL_Type.Base_URL, "/KronosUltimate/rest-1.v1/Data/Task/" + taskIDNumber));
        final JSONObject task = new JSONObject(response.getBody());
        extractInformationOfTask(task, targetDirectory);
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
                runExtractionOnEpicData("example");
                runExtractionOnStoryData("example");
                runExtractionOnDefectData("example");
                runExtractionOnTaskData("example");

        } catch (/*KeyManagementException | NoSuchAlgorithmException | KeyStoreException | IOException |*/ Exception e) {
            e.printStackTrace();
        }
    }
}