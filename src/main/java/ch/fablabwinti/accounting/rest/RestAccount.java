package ch.fablabwinti.accounting.rest;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import com.fasterxml.jackson.annotation.JsonTypeName;

import java.util.List;
import java.util.Map;

@JsonIgnoreProperties(ignoreUnknown = true)


public class RestAccount extends RestObject {

}

//public class RestAccount {
//
//        // @JsonProperty("wrapper")
//
//        //public RestAccount() {
//        //       properties = new Properties();
//        //}
//
//        @JsonIgnoreProperties(ignoreUnknown = true)
//        static class Properties {
//                String title;
//
//                Double amount;
//
//                Double openingentry;
//
//                public String getTitle() {
//                        return title;
//                }
//
//                public void setTitle(String title) {
//                        this.title = title;
//                }
///*
//                public Double getAmount() {
//                        return amount;
//                }
//
//                public void setAmount(Double amount) {
//                        this.amount = amount;
//                }
//*/
//                public Double getOpeningentry() {
//                        return openingentry;
//                }
//
//                public void setOpeningentry(Double openingentry) {
//                        this.openingentry = openingentry;
//                }
//        }
//
//        private String type;
//
//        private boolean readonly;
//
//        private Properties properties;
//
//        private Map<String, String> children;
//
//        private Map<String, int[]> links;
//
//        private int[] parents;
//
//        public String getType() {
//                return type;
//        }
//
//        public void setType(String type) {
//                this.type = type;
//        }
//
//        public boolean isReadonly() {
//                return readonly;
//        }
//
//        public void setReadonly(boolean readonly) {
//                this.readonly = readonly;
//        }
//
//        public Properties getProperties() {
//                return properties;
//        }
//
//        public void setProperties(Properties properties) {
//                this.properties = properties;
//        }
//
//        public Map<String, String> getChildren() {
//                return children;
//        }
//
//        public void setChildren(Map<String, String> children) {
//                this.children = children;
//        }
//
//        public Map<String, int[]> getLinks() {
//                return links;
//        }
//
//        public void setLinks(Map<String, int[]> links) {
//                this.links = links;
//        }
//
//        public int[] getParents() {
//                return parents;
//        }
//
//        public void setParents(int[] parents) {
//                this.parents = parents;
//        }
//}
