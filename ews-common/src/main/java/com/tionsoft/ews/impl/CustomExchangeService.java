package com.tionsoft.ews.impl;

import java.security.GeneralSecurityException;

import microsoft.exchange.webservices.data.EWSConstants;
import microsoft.exchange.webservices.data.core.EwsSSLProtocolSocketFactory;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;

import org.apache.http.config.Registry;
import org.apache.http.config.RegistryBuilder;
import org.apache.http.conn.socket.ConnectionSocketFactory;
import org.apache.http.conn.socket.PlainConnectionSocketFactory;
import org.apache.http.conn.ssl.NoopHostnameVerifier;

public class CustomExchangeService extends ExchangeService {
	 public CustomExchangeService(ExchangeVersion exchange2010Sp2) {
		 super(exchange2010Sp2);
		// TODO Auto-generated constructor stub
	}

	@Override
	  protected Registry<ConnectionSocketFactory> createConnectionSocketFactoryRegistry() {
	    try {
	      return RegistryBuilder.<ConnectionSocketFactory>create()
	          .register(EWSConstants.HTTP_SCHEME, new PlainConnectionSocketFactory())
	          .register(EWSConstants.HTTPS_SCHEME, EwsSSLProtocolSocketFactory.build(
	              null, NoopHostnameVerifier.INSTANCE
	          ))
	          .build();
	    } catch (GeneralSecurityException e) {
	      throw new RuntimeException(
	          "Could not initialize ConnectionSocketFactory instances for HttpClientConnectionManager", e
	      );
	    }
	  }
}
