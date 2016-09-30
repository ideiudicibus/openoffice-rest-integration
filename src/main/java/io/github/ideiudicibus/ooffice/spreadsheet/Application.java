package io.github.ideiudicibus.ooffice.spreadsheet;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.context.ApplicationContext;

@SpringBootApplication(scanBasePackages = "io.github.ideiudicibus.ooffice.spreadsheet")

public class Application {

    public static void main(String[] args) {
        SpringApplication.run(Application.class, args);
        //ApplicationContext ctx = new SpringApplicationBuilder(ObjectPoolConfig.class).run(args);
    }
}
