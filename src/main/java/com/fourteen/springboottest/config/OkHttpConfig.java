package com.fourteen.springboottest.config;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import okhttp3.*;
import okhttp3.internal.connection.RouteException;
import okhttp3.internal.http.RealInterceptorChain;
import okhttp3.internal.http2.ConnectionShutdownException;
import org.apache.commons.lang3.StringUtils;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.Primary;

import javax.net.ssl.SSLHandshakeException;
import javax.net.ssl.SSLPeerUnverifiedException;
import java.io.IOException;
import java.io.InterruptedIOException;
import java.net.ProtocolException;
import java.net.SocketTimeoutException;
import java.security.cert.CertificateException;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;


/**
 * @author huangwei6
 * @since 2022/8/18
 */
@Slf4j
@Configuration
@RequiredArgsConstructor
public class OkHttpConfig {


    @Bean
    @Primary
    public OkHttpClient buildOkHttpClient() {
        OkHttpClient okHttpClient = new OkHttpClient.Builder()
                .connectionPool(new ConnectionPool(50, 10, TimeUnit.MINUTES))
                .connectTimeout(10, TimeUnit.SECONDS)
                .readTimeout(10, TimeUnit.SECONDS)
                .writeTimeout(10, TimeUnit.SECONDS)
                .pingInterval(1, TimeUnit.MINUTES)
                .retryOnConnectionFailure(true)
                .addInterceptor(this::retryInterceptor)
                .addInterceptor(this::traceIdInterceptor)
                .build();
        log.info("OkHttpClient build success");
        return okHttpClient;
    }

    public Response traceIdInterceptor(Interceptor.Chain chain) throws IOException {
        Request oldReq = chain.request();
        Request.Builder builder = oldReq.newBuilder();
        builder.method(oldReq.method(), oldReq.body());
        builder.headers(oldReq.headers());
        Request newReq = builder.build();
        return chain.proceed(newReq);
    }

    public Response retryInterceptor(Interceptor.Chain chain) {
        Request request = chain.request();
        RealInterceptorChain realChain = (RealInterceptorChain) chain;
        AtomicInteger count = new AtomicInteger();
        boolean retry = true;
        Response response = null;
        while (retry) {
            try {
                response = realChain.proceed(request);
                retry = false;
            } catch (RouteException e) {
                if (recover(e.getLastConnectException(), false, request)) {
                    log.error("{}, Do not retry, Route Error {}", request, e.getLastConnectException().getMessage());
                    return error(request, e.getLastConnectException().getMessage());
                }
                sleep(request, count);
            } catch (NullPointerException | IllegalStateException e) {
                sleep(request, count);
            } catch (IOException e) {
                boolean requestSendStarted = !(e instanceof ConnectionShutdownException);
                if (this.recover(e, requestSendStarted, request)) {
                    log.error("{}, Do not retry, IO Error {}", request, e.getMessage());
                    return error(request, e.getMessage());
                }
                sleep(request, count);
            } finally {
                if (count.get() > 5) {
                    log.info("{}, Do not retry, Over max retry limit", request);
                    retry = false;
                }
            }
        }
        return response;
    }

    private boolean recover(IOException e, boolean requestSendStarted, Request userRequest) {
        if (true) {
            if (requestSendStarted && StringUtils.containsIgnoreCase(e.getMessage(), "timeout")) {
                return !StringUtils.equalsIgnoreCase(userRequest.method(), "GET");
            } else {
                return !isRecoverable(e, requestSendStarted);
            }
        }
        return true;
    }

    private boolean isRecoverable(IOException e, boolean requestSendStarted) {
        if (e instanceof ProtocolException) {
            return false;
        }
        if (e instanceof InterruptedIOException) {
            return e instanceof SocketTimeoutException && !requestSendStarted;
        }
        if (e instanceof SSLHandshakeException) {
            if (e.getCause() instanceof CertificateException) {
                return false;
            }
        }
        return !(e instanceof SSLPeerUnverifiedException);
    }

    private void sleep(Request request, AtomicInteger count) {
        long seconds = 5;
        try {
            count.incrementAndGet();
            log.info("{}, Start the {}th retry after {}s", request.toString(), count.get(), seconds);
            TimeUnit.SECONDS.sleep(seconds);
        } catch (InterruptedException ignored) {
        }
    }

    private Response error(Request request, String message) {
        Response.Builder builder = new Response.Builder();
        builder.code(500);
        builder.message(message);
        builder.request(request);
        builder.protocol(Protocol.HTTP_1_1);
        return builder.build();
    }
}
