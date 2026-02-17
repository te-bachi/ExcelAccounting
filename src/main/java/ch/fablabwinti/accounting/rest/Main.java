package ch.fablabwinti.accounting.rest;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.util.*;

public class Main {

    public static void main(String[] args) throws Exception {
        //HttpTest test = new HttpTest();
        //RestAccount restAccountJackson = test.syncJackson();
        //GenericTest<RestAccount>  test = new GenericTest<>();
        //RestAccount restAccountJackson = test.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/account/359"), RestAccount.class);


        RestClient restClient = new RestClient();


        System.out.println("=== Periodgroup (Buchhaltung) ===================================================");
        RestObjectList restPeriodgroupObjectList = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/periodgroup"), RestObjectList.class);
        List<RestPeriodgroup> restPeriodgroupList = new ArrayList<>();
        try {
            for (int periodgroupId : restPeriodgroupObjectList.getObjects()) {
                RestPeriodgroup restPeriodgroup = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/periodgroup/" + periodgroupId), RestPeriodgroup.class);
                restPeriodgroupList.add(restPeriodgroup);
            }
        } catch (NullPointerException e) {
            System.out.println("error: " + restPeriodgroupObjectList.getError());
        }

        System.out.println("=== Periodchain (Buchhaltungskette: 2024 - 2025 mit AccoutGroupTemplates) =======");
        RestObjectList restPeriodchainObjectList = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/periodchain"), RestObjectList.class);
        List<RestPeriodchain> restPeriodchainList = new ArrayList<>();
        for (int periodchainId : restPeriodchainObjectList.getObjects()) {
            RestPeriodchain restPeriodchain = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/periodchain/" + periodchainId), RestPeriodchain.class);
            restPeriodchainList.add(restPeriodchain);
        }

        System.out.println("=== Period List (Liste von Perioden: 2024, 2025, etc.) ==========================");
        RestObjectList restPeriodObjectList = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/period"), RestObjectList.class);
        List<RestPeriod> restPeriodList = new ArrayList<>();
        for (int periodId : restPeriodObjectList.getObjects()) {
            RestPeriod restPeriod = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/period/" + periodId), RestPeriod.class);
            /* nur 2025 ohne 2024, sonst spuckt der REST Client null zurück */
            if (restPeriod.getProperties().get("title").equals("Buchungsperiode 2024")) {
                System.out.println("<skipping Buchungsperiode 2024>");
                continue;
            } else if (restPeriod.getProperties().get("title").equals("Buchungsperiode 2025")) {
                System.out.println("<go-on with Buchungsperiode 2025>");
            } else {
                System.out.println("<skipping unknow Buchungsperiode " + restPeriod.getProperties().get("title") +  ">");
            }
            restPeriod.setId(periodId);
            restPeriodList.add(restPeriod);
            /* keine Buchungen suchen, sonst spuckt der REST Client null zurück
            System.out.println("  --- Entry Group List (Buchungen) -----------------------------");
            for (int entryGroupId : restPeriod.getChildren().get("entrygroup")) {
                RestEntrygroup restEntryGroup = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/entrygroup/" + entryGroupId), RestEntrygroup.class);
                restEntryGroup.setId(entryGroupId);
                restPeriod.getEntryGroupList().add(restEntryGroup);
            }
            */
            System.out.println("  --- Account Group List (Kontengruppe) ---------------------------");
            for (int accountGroupId : restPeriod.getChildren().get("accountgroup")) {
                RestAccountGroup restAccountGroup = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/accountgroup/" + accountGroupId), RestAccountGroup.class);
                restAccountGroup.setId(accountGroupId);
                restPeriod.getAccountGroupList().add(restAccountGroup);
                System.out.println("    --- Account List (Konto) -------------------------------");
                try {
                    for (int accountId : restAccountGroup.getChildren().get("account")) {
                        RestAccount restAccount = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/account/" + accountId), RestAccount.class);
                        restAccount.setId(accountId);
                        restAccountGroup.getAccountList().add(restAccount);
                    }
                } catch (NullPointerException e) {
                    System.out.println("<skipping>");
                }
            }

            //for ()
        }
        Map<String, String> idMap = new LinkedHashMap<String, String>();
        Map<String, String> nameMap = new LinkedHashMap<String, String>();

        for (RestPeriod period : restPeriodList) {
            System.out.println("==> " + period.getId() + ": " + period.getProperties().get("title"));
            /* keine Buchungen printed, sonst spuckt der REST Client null zurück
            System.out.println("  --> entrygroup");
            for (RestEntrygroup entryGroup : period.getEntryGroupList()) {
                System.out.println("  --> " + entryGroup.getId() + ": " + entryGroup.getProperties().get("title"));
            }
            */
            System.out.println("  --> accountgroup");
            for (RestAccountGroup accountGroup : period.getAccountGroupList()) {
                System.out.println("  --> " + accountGroup.getId() + ": " + accountGroup.getProperties().get("title"));
                System.out.println("    --> account");
                for (RestAccount account : accountGroup.getAccountList()) {
                    System.out.println("    --> " + account.getId() + ": " + account.getProperties().get("title"));
                    String idName[] = account.getProperties().get("title").split(" ", 2);
                    idMap.put(idName[0], String.valueOf(account.getId()));
                    nameMap.put(idName[0], idName[1]);
                }
            }
        }

        try {
            Properties idProperties = new Properties();
            idProperties.putAll(idMap);
            idProperties.store(new FileOutputStream("id.properties"), "id");
            Properties nameProperties = new Properties();
            nameProperties.putAll(nameMap);
            nameProperties.store(new FileOutputStream("name.properties"), "name");
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }

        System.out.println("=== Accountgrouptemplate ===========================");
        RestObjectList restAccountgrouptemplateObjectList = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/accountgrouptemplate"), RestObjectList.class);
        List<RestObject> restAccountgrouptemplateList = new ArrayList<>();
        for (int accountgrouptemplateId : restAccountgrouptemplateObjectList.getObjects()) {
            RestObject restAccountgrouptemplate = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/accountgrouptemplate/" + accountgrouptemplateId), RestObject.class);
            restAccountgrouptemplate.setId(accountgrouptemplateId);
            restAccountgrouptemplateList.add(restAccountgrouptemplate);
        }

        /*
        for (RestObject account : restAccountgrouptemplateList) {
            System.out.println("==> " + account.getId() + ": " + account.getProperties().get("title"));
        }
        */


        /*
        RestPeriodchain restPeriodchain = new RestPeriodchain();
        restPeriodchain.setProperties(Stream.of(new String[][] {
                { "title", "Kontenrahmen Bla" },
                { "sourcechart", "custom" }
        }).collect(Collectors.toMap(data -> data[0], data -> data[1])));


        restPeriodchain.setCh(Stream.of(new Object[][] {
                { "title", "Kontenrahmen Bla" },
                { "sourcechart", "custom" }
        }).collect(Collectors.toMap(data -> data[0], data -> data[1])));
        */

        //RestAccount restAccountGson = test.syncGson();
        //restAccountJackson.getProperties().setTitle("bla bla");
        //test.postGson(restAccountJackson);

        System.out.println("Done!");
    }

    /*
    String getRequest() throws Exception {
        HttpRequest request = HttpRequest.newBuilder()
                .uri(new URI("https://fablabwinti.webling.ch/api/1/account/359"))
                .header("apikey", "***masked***")
                .GET()
                .build();

        HttpClient client = HttpClient.newBuilder()
                .build();

        HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

        String body = response.body();

        System.out.println(body);

        return body;
    }

    String postRequest(String jsonString) throws Exception {

        System.out.println(jsonString);

        HttpRequest request = HttpRequest.newBuilder()
                .uri(new URI("https://fablabwinti.webling.ch/api/1/account"))
                .header("apikey", "***masked***")
                .POST(HttpRequest.BodyPublishers.ofString(jsonString))
                .build();

        HttpClient client = HttpClient.newBuilder()
                .build();

        HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

        System.out.println("statusCode=" + response.statusCode());
        return response.body();
    }

    public <T> T syncJackson() throws Exception {
        ObjectMapper objectMapper = new ObjectMapper();
        String response = getRequest();
        TypeReference<T> typeRef = new TypeReference<T>() {};
        T t = objectMapper.readValue(response, typeRef);

        return t;
    }

    public RestAccount syncGson() throws Exception {
        Gson gson = new GsonBuilder().create();
        String response = getRequest();
        RestAccount restAccount = gson.fromJson(response, new TypeToken<RestAccount>(){}.getType());
        return restAccount;
    }

    public void postGson(RestAccount restAccount) throws Exception {
        Gson gson = new GsonBuilder().create();
        String jsonString = gson.toJson(restAccount);

        String response = postRequest(jsonString);

    }

    */
}
