package ch.fablabwinti.accounting.rest;

import com.fasterxml.jackson.databind.ObjectMapper;

import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;

public class RestClient {

    private HttpClient client;

    public RestClient() {
        this.client = HttpClient.newBuilder().build();
    }

    String getRequest(URI uri) throws Exception {
        HttpRequest request = HttpRequest.newBuilder()
                .uri(uri)
                .header("apikey", "7be1772dfbd171d1f675aaee19a5c9b3")
                .GET()
                .build();

        HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

        //String body = "{\"result\":" + response.body() + "}";
        String body = response.body();

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

    public <T> T syncJackson(URI uri, Class<T> klass) throws Exception {
        ObjectMapper objectMapper = new ObjectMapper();
        String response = getRequest(uri);

        /*
        Object json = objectMapper.readValue(response, Object.class);
        String prettyPrint = objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(json);
        System.out.println(prettyPrint);
         */

        T resp = objectMapper.readValue(response, klass);
        String prettyPrintResp = objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(resp);
        System.out.println(prettyPrintResp);

        return resp;
    }

}
