package io.muenchendigital.digiwf.outlookintegration.service;

import lombok.SneakyThrows;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.availability.AvailabilityData;
import microsoft.exchange.webservices.data.core.enumeration.availability.FreeBusyViewType;
import microsoft.exchange.webservices.data.core.enumeration.availability.MeetingAttendeeType;
import microsoft.exchange.webservices.data.core.enumeration.availability.SuggestionQuality;
import microsoft.exchange.webservices.data.core.enumeration.service.SendInvitationsMode;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.misc.availability.AttendeeInfo;
import microsoft.exchange.webservices.data.misc.availability.AvailabilityOptions;
import microsoft.exchange.webservices.data.misc.availability.GetUserAvailabilityResults;
import microsoft.exchange.webservices.data.misc.availability.TimeWindow;
import microsoft.exchange.webservices.data.property.complex.Attendee;
import microsoft.exchange.webservices.data.property.complex.availability.Suggestion;
import microsoft.exchange.webservices.data.property.complex.availability.TimeSuggestion;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.net.URI;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

@Service
public class OutlookService {

    private ExchangeService exchangeService;

    public OutlookService(
            @Value("${exchange.username}") String exchangeUsername,
            @Value("${exchange.password}") String exchangePassword,
            @Value("${exchange.uri}") URI exchangeUri
    ) {
        exchangeService = new ExchangeService();
        WebCredentials webCredentials = new WebCredentials(exchangeUsername, exchangePassword);
        exchangeService.setCredentials(webCredentials);
        exchangeService.setUrl(exchangeUri);
    }

    @SneakyThrows
    public void getAvailableTimesForUser(String mail, int durationMinutes, TimeWindow timeWindow) {
        ArrayList<AttendeeInfo> ais = new ArrayList<>();
        AttendeeInfo ai = new AttendeeInfo(mail, MeetingAttendeeType.Required, true);
        ais.add(ai);
        AvailabilityOptions ao = new AvailabilityOptions();
        ao.setMinimumSuggestionQuality(SuggestionQuality.Good);
        ao.setMeetingDuration(durationMinutes);
        ao.setRequestedFreeBusyView(FreeBusyViewType.FreeBusy);
        ao.setDetailedSuggestionsWindow(timeWindow);
        GetUserAvailabilityResults userAvailability = exchangeService.getUserAvailability(ais, timeWindow, AvailabilityData.FreeBusyAndSuggestions, ao);
        for (Suggestion suggestion : userAvailability.getSuggestionsResponse().getSuggestions()) {
            for (TimeSuggestion timeSuggestion : suggestion.getTimeSuggestions()) {
                System.out.println(timeSuggestion.getMeetingTime());
            }
        }
        return;
    }

    @SneakyThrows
    public void createAppointment(String mail) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date start = simpleDateFormat.parse("2022-06-16 12:00:00");
        Date end = simpleDateFormat.parse("2022-06-16 13:00:00");
        Appointment appointment = new Appointment(exchangeService);
        appointment.setSubject("Test");
        appointment.setStart(start);
        appointment.setEnd(end);
        appointment.getRequiredAttendees().add(new Attendee(mail));
        appointment.save(SendInvitationsMode.SendOnlyToAll);
    }

}
