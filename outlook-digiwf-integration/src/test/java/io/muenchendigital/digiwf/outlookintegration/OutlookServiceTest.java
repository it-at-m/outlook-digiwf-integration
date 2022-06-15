package io.muenchendigital.digiwf.outlookintegration;

import io.muenchendigital.digiwf.outlookintegration.service.OutlookService;
import lombok.SneakyThrows;
import microsoft.exchange.webservices.data.misc.availability.TimeWindow;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.TestPropertySource;

import java.text.SimpleDateFormat;
import java.util.Date;

@SpringBootTest
@ContextConfiguration(classes = Application.class)
@TestPropertySource(locations="classpath:application.properties")
public class OutlookServiceTest {

    @Autowired
    private OutlookService outlookService;

    @SneakyThrows
    @Test
    public void test() {
        outlookService.createAppointment("testuser2.ews@testlhm.de");
        SimpleDateFormat f = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date start = f.parse("2022-06-16 00:00:00");
        Date end = f.parse("2022-06-16 23:59:59");
        outlookService.getAvailableTimesForUser("testuser2.ews@testlhm.de", 60, new TimeWindow(start, end));
    }

}
