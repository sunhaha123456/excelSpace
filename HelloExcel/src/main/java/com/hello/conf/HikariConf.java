package com.hello.conf;

import com.zaxxer.hikari.HikariDataSource;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import javax.sql.DataSource;

@Configuration
public class HikariConf {

	@Bean(name = "primaryDataSource", destroyMethod = "close")
	@Qualifier("primaryDataSource")
	@ConfigurationProperties(prefix = "hikari")
	public DataSource dataSource() {
		return new HikariDataSource();
	}

}