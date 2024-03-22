package ch.fablabwinti.accounting.rest;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.reflect.TypeToken;

import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class Main {

    public static void main(String[] args) throws Exception {
        //HttpTest test = new HttpTest();
        //RestAccount restAccountJackson = test.syncJackson();
        //GenericTest<RestAccount>  test = new GenericTest<>();
        //RestAccount restAccountJackson = test.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/account/359"), RestAccount.class);


        RestClient restClient = new RestClient();

        System.out.println("=== Periodgroup ====================================");
        RestObjecList restPeriodgroupObjectList = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/periodgroup"), RestObjecList.class);
        List<RestPeriodgroup> restPeriodgroupList = new ArrayList<>();
        for (int periodgroupId : restPeriodgroupObjectList.getObjects()) {
            RestPeriodgroup restPeriodgroup = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/periodgroup/" + periodgroupId), RestPeriodgroup.class);
            restPeriodgroupList.add(restPeriodgroup);
        }

        System.out.println("=== Periodchain ====================================");
        RestObjecList restPeriodchainObjectList = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/periodchain"), RestObjecList.class);
        List<RestPeriodchain> restPeriodchainList = new ArrayList<>();
        for (int periodchainId : restPeriodchainObjectList.getObjects()) {
            RestPeriodchain restPeriodchain = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/periodchain/" + periodchainId), RestPeriodchain.class);
            restPeriodchainList.add(restPeriodchain);
        }

        System.out.println("=== Period =========================================");
        RestObjecList restPeriodObjectList = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/period"), RestObjecList.class);
        List<RestPeriod> restPeriodList = new ArrayList<>();
        for (int periodId : restPeriodObjectList.getObjects()) {
            RestPeriod restPeriod = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/period/" + periodId), RestPeriod.class);
            restPeriodList.add(restPeriod);
        }

        System.out.println("=== Accountgrouptemplate ===========================");
        RestObjecList restAccountgrouptemplateObjectList = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/accountgrouptemplate"), RestObjecList.class);
        List<RestAccountgrouptemplate> restAccountgrouptemplateList = new ArrayList<>();
        for (int accountgrouptemplateId : restAccountgrouptemplateObjectList.getObjects()) {
            RestAccountgrouptemplate restAccountgrouptemplate = restClient.syncJackson(new URI("https://fablabwinti.webling.ch/api/1/accountgrouptemplate/" + accountgrouptemplateId), RestAccountgrouptemplate.class);
            restAccountgrouptemplateList.add(restAccountgrouptemplate);
        }



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

    String getRequest() throws Exception {
        HttpRequest request = HttpRequest.newBuilder()
                .uri(new URI("https://fablabwinti.webling.ch/api/1/account/359"))
                .header("apikey", "7be1772dfbd171d1f675aaee19a5c9b3")
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
                .header("apikey", "7be1772dfbd171d1f675aaee19a5c9b3")
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
}
