package com.tionsoft.ews;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.inject.Inject;

import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.item.Appointment;

import org.apache.commons.lang3.time.DateUtils;
import org.junit.Assert;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations="/spring/test-application.xml")
public class EWSServiceTest {

	@Inject
	private EWSService service;
	
	@Test
	public void testCreateSchedule() {
		Assert.assertTrue(this.createSchedule("test subject"));
	}
	
	@Test
	public void testGetTodayScheduleCount() {
		int count = this.getTodayScheduleCount();
		Assert.assertTrue(count >= 0);
	}
	
	@Test
	public void testSearchSchedule() {
		int count = this.searchCalendar("test subject");
		Assert.assertTrue(count >= 0);
	}
	
	@Test
	public void testDeleteSchdule() {
		Assert.assertTrue(this.deleteSchedule("test subject"));
	}
	
	public boolean createSchedule(String subject) {
		Date startDate = new Date();
		Date endDate = DateUtils.addMinutes(startDate, 60);
		List<String> attendees = new ArrayList<String>();
		attendees.add("neoneox@hanmail.net");
		
		try {
			String uniqueId = service.createSchedule(subject, "body test", startDate, endDate, -1, 60, "location test", "", null, null, null, null, 0);
//			System.out.println("create appointment success, id " + uniqueId);
			return true;
		} catch (EWSServiceException e) {
			e.printStackTrace();
			return false;
		}
	}
	
	public int getTodayScheduleCount() {
		try {
			return service.getTodayScheduleCount();
		} catch (EWSServiceException e) {
			e.printStackTrace();
			return -1;
		}
	}
	
	public int searchCalendar(String searchWord) {
		Date now = new Date();
		Date startTime = DateUtils.addDays(now, -1);
		Date endTime = DateUtils.addDays(now, 30);
		ArrayList<Appointment> results;
		try {
			results = service.getScheduleSearchPageList(startTime, endTime, 1, searchWord, 20, 1);
			return results.size();
		} catch (Exception e) {
			e.printStackTrace();
			return -1;
		}
	}
	
	public boolean deleteSchedule(String subject) {
		Date now = new Date();
		Date startTime = DateUtils.addDays(now, -30);
		Date endTime = DateUtils.addDays(now, 30);
		try {
			ArrayList<Appointment> results = service.getScheduleList(startTime, endTime, 0);
		    for (Appointment appt : results) {
//		        System.out.println("SUBJECT=====" + appt.getSubject());
//		        System.out.println("BODY========" + MessageBody.getStringFromMessageBody(appt.getBody()));
		        if (appt.getSubject().equals(subject)) {
//		        	System.out.println("delete subject : " + subject);
		        	service.deleteSchedule(appt);
		        }
		    }
		    return true;
		} catch (EWSServiceException e) {
			e.printStackTrace();
			return false;
		} catch (ServiceLocalException e) {
			e.printStackTrace();
			return false;
		}
	}
}
