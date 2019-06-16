package com.test.word.wordpdf;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.web.support.SpringBootServletInitializer;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.Configuration;

@Configuration
@SpringBootApplication
@ComponentScan(basePackages = {
        "com.test.word.wordpdf.*"
})
public class WordTestApplication extends SpringBootServletInitializer {
    public static void main(String[] args) {
        SpringApplication.run(WordTestApplication.class, args);
    }

    protected SpringApplicationBuilder config(SpringApplicationBuilder applicationBuilder) {
        return applicationBuilder.sources(WordTestApplication.class);
    }
}