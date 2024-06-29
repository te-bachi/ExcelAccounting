package ch.fablabwinti.accounting.rest;

import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.FileInputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URL;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.Properties;

public class RestClient {

    private HttpClient client;
    private String apikey;

    public RestClient() throws IOException {
        this.client = HttpClient.newBuilder().build();

        Properties appProps = new Properties();
        URL url = RestClient.class.getResource("apikey.properties");
        appProps.load(url.openStream());
        this.apikey = appProps.getProperty("apikey");
    }

    String getRequest(URI uri) throws Exception {
        HttpRequest request = HttpRequest.newBuilder()
                .uri(uri)
                .header("apikey", this.apikey)
                .GET()
                .build();

        HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

        //String body = "{\"result\":" + response.body() + "}";
        String body = response.body();

        return body;
    }

    Pair<Integer, String> postRequest(URI uri, String jsonString) throws Exception {

        //System.out.println(jsonString);

        HttpRequest request = HttpRequest.newBuilder()
                .uri(uri)
                .header("apikey", this.apikey)
                .POST(HttpRequest.BodyPublishers.ofString(jsonString))
                .build();

        HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

        System.out.println("statusCode=" + response.statusCode());
        String body = response.body();

        Pair<Integer, String> pair = new Pair(Integer.valueOf(response.statusCode()), response.body());
        return pair;
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
        //System.out.println(prettyPrintResp);

        return resp;
    }

    public <T> Pair<Integer, String> postJackson(URI uri, Class<T> klass) throws Exception {
        ObjectMapper objectMapper = new ObjectMapper();

        String json = objectMapper.writeValueAsString(klass);
        Pair<Integer, String> pair = postRequest(uri, json);

        return pair;
    }

}
