package com.tionsoft.ews;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import microsoft.exchange.webservices.data.core.enumeration.property.Sensitivity;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.property.complex.AttachmentCollection;
import microsoft.exchange.webservices.data.property.complex.Attendee;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.property.complex.MessageBody;

import org.apache.commons.lang3.time.DateUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.support.ClassPathXmlApplicationContext;

public class EWSTest {

	@Autowired
	private EWSService service;

	public static void main(String[] args) {
		try (final ClassPathXmlApplicationContext context = new ClassPathXmlApplicationContext(
				"/spring/application.xml")) {
			context.registerShutdownHook();
			EWSTest test = (EWSTest)context.getBean("testBean");
			test.craeteAppointment("subject test");
			test.getTodayScheduleCount();
			test.findAppointments();
			test.searchCalendar("test");
			test.updateAppointment("subject test");
			test.deleteAppointment("subject test");
		} catch (final Exception ex) {
			ex.printStackTrace();
		}
	}

	public void craeteAppointment(String subject) {
		Date startDate = new Date();
		Date endDate = DateUtils.addMinutes(startDate, 60);
		List<String> attendees = new ArrayList<String>();
		attendees.add("test@test.com");
		try {
			String uniqueId = service.createSchedule(subject, "body test", startDate, endDate, -1, 60, "location test", "", null, null, null, null, 0);
			System.out.println("create appointment success, id " + uniqueId);
		} catch (EWSServiceException e) {
			e.printStackTrace();
		}
	}
	
	public void getTodayScheduleCount() {
		try {
			int count = service.getTodayScheduleCount();
			System.out.println("today schedule count : " + count);
		} catch (EWSServiceException e) {
			e.printStackTrace();
		}
	}
	
	public void findAppointments() {
		Date now = new Date();
		Date startTime = DateUtils.addDays(now, -1);
		Date endTime = DateUtils.addDays(now, 30);
		ArrayList<Appointment> results;
		try {
			results = service.getScheduleList(startTime, endTime, 0);
		    for (Appointment appt : results) {
		        System.out.println("SUBJECT=====" + appt.getSubject());
		        System.out.println("BODY========" + MessageBody.getStringFromMessageBody(appt.getBody()));
		        if (appt.getHasAttachments()) { 
		            AttachmentCollection attachmentsCol = appt.getAttachments(); 
		            for (int i = 0; i < attachmentsCol.getCount(); i++) {
		                FileAttachment attachment = (FileAttachment)attachmentsCol.getPropertyAtIndex(i);
		                String where = "/Users/neox/Downloads/ews/" + attachment.getName();
		                System.out.println(i + ") " + where + ", content id : " + attachment.getContentId());
		                attachment.load(where);
		            }
		        }
		    }
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void searchCalendar(String searchWord) {
		Date now = new Date();
		Date startTime = DateUtils.addDays(now, -1);
		Date endTime = DateUtils.addDays(now, 30);
		ArrayList<Appointment> results;
		try {
			results = service.getScheduleSearchPageList(startTime, endTime, 1, searchWord, 20, 1);
		    for (Appointment appt : results) {
		        System.out.println("SUBJECT=====" + appt.getSubject());
		        System.out.println("BODY========" + MessageBody.getStringFromMessageBody(appt.getBody()));
		        if (appt.getHasAttachments()) { 
		            AttachmentCollection attachmentsCol = appt.getAttachments(); 
		            for (int i = 0; i < attachmentsCol.getCount(); i++) {
		                FileAttachment attachment = (FileAttachment)attachmentsCol.getPropertyAtIndex(i);
		                String where = "/Users/neox/Downloads/ews/" + attachment.getName();
		                System.out.println(i + ") " + where + ", content id : " + attachment.getContentId());
		                attachment.load(where);
		            }
		        }
		    }
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void deleteAppointment(String subject) {
		Date now = new Date();
		Date startTime = DateUtils.addDays(now, -30);
		Date endTime = DateUtils.addDays(now, 30);
		try {
			ArrayList<Appointment> results = service.getScheduleList(startTime, endTime, 0);
		    for (Appointment appt : results) {
//		        System.out.println("SUBJECT=====" + appt.getSubject());
		        if (appt.getSubject().equals(subject)) {
		        	System.out.println("delete subject : " + subject);
		        	service.deleteSchedule(appt);
		        }
		    }
		} catch (EWSServiceException e) {
			e.printStackTrace();
		} catch (ServiceLocalException e) {
			e.printStackTrace();
		}
	}
	
	public String getAddressFromAttendee(Attendee attendee) {
		return attendee.getAddress();
	}
	
	public List<String> translateAttendeeList(List<Attendee> attendeeList) {
		List<String> translatedList = new ArrayList<>();
		for (Attendee attendee : attendeeList) {
			translatedList.add(getAddressFromAttendee(attendee));
		}
		return translatedList;
	}
	
	public void updateAppointment(String subject) {
		try {
			Date now = new Date();
			Date startTime = DateUtils.addDays(now, -30);
			Date endTime = DateUtils.addDays(now, 30);
			ArrayList<Appointment> results = service.getScheduleList(startTime, endTime, 0);
			for (Appointment appt : results) {
				//		        System.out.println("SUBJECT=====" + appt.getSubject());
				if (appt.getSubject().equals(subject)) {
					System.out.println("update subject : " + subject);
					service.modifyAppointment(appt.getId(), "update subject", "update body", appt.getStart(), appt.getEnd(), 
							appt.getIsAllDayEvent() ? 1 : 0, 
									appt.getIsReminderSet() ? appt.getReminderMinutesBeforeStart() : -1, 
											appt.getLocation(), appt.getCategories().toString(), 
											translateAttendeeList(appt.getRequiredAttendees().getItems()), 
											translateAttendeeList(appt.getOptionalAttendees().getItems()), 
											translateAttendeeList(appt.getResources().getItems()), 
											appt.getRecurrence(), appt.getSensitivity() == Sensitivity.Private ? 0 : 1);
				}
			}
		} catch (ServiceLocalException | EWSServiceException e) {
			e.printStackTrace();
		}
	}
}
