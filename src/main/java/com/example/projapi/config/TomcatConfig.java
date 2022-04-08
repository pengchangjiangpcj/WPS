package com.example.projapi.config;

import org.apache.catalina.connector.Connector;
import org.springframework.boot.web.embedded.tomcat.TomcatServletWebServerFactory;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * @description: tomcat配置特殊字符
 * @author: PCJ
 * @create: 2021-09-29 10:33
 **/
@Configuration
public class TomcatConfig {

    @Bean
    public TomcatServletWebServerFactory webServerFactory(){
        TomcatServletWebServerFactory factory = new TomcatServletWebServerFactory();
        factory.addConnectorCustomizers((Connector connector) -> {
            connector.setProperty("relaxedPathChars", "\"#<>[\\]^`{|}");
            connector.setProperty("relaxedQueryChars", "\"#<>[\\]^`{|}");
        });
        return factory;
    }

}
