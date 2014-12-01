/**
 * 
 */
package com.microsoft.exchange.config;

import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.EnableAspectJAutoProxy;
import org.springframework.context.annotation.Import;
import org.springframework.retry.annotation.EnableRetry;

/**
 * @author ctcudd
 *
 */
@Configuration
@EnableAspectJAutoProxy(proxyTargetClass = true)
@EnableRetry
@Import(value=ImpersonationConfig.class)
public class TestConfig {

}
