package com.tionsoft.ews.impl;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.property.Sensitivity;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.enumeration.service.SendInvitationsMode;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.AppointmentSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.AttachmentCollection;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.property.complex.StringList;
import microsoft.exchange.webservices.data.property.complex.recurrence.pattern.Recurrence;
import microsoft.exchange.webservices.data.search.CalendarView;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

import org.joda.time.DateTime;

import com.tionsoft.ews.EWSService;
import com.tionsoft.ews.EWSServiceException;

public class EWSServiceImpl implements EWSService {

	private static ExchangeService service;
	
	private String id = "";
	private String password = "";
	private String uri = "";

	/**
	 * Firstly check, whether "https://webmail.xxxx.com/ews/Services.wsdl" and "https://webmail.xxxx.com/ews/Exchange.asmx"
	 * is accessible, if yes that means the Exchange Webservice is enabled on your MS Exchange.
	 */

	static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
		public boolean autodiscoverRedirectionUrlValidationCallback(
				String redirectionUrl) {
			System.out.println("callback called!!!");
			return redirectionUrl.toLowerCase().startsWith("https://");
		}
	}

	/**
	 * Initialize the Exchange Credentials.
	 * Don't forget to replace the "USRNAME","PWD","DOMAIN_NAME" variables.  
	 */
	public EWSServiceImpl(String id, String password, String uri) {
		this.id = id;
		this.password = password;
		this.uri = uri;
		
		service = new CustomExchangeService(ExchangeVersion.Exchange2010_SP2);
		//service = new ExchangeService(ExchangeVersion.Exchange2007_SP1); //depending on the version of your Exchange. 
		try {
			service.setUrl(new URI(this.uri));
			ExchangeCredentials credentials = new WebCredentials(this.id, this.password);
			service.setCredentials(credentials);
		} catch (URISyntaxException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * get today schedule count 
	 */
	@Override
	public int getTodayScheduleCount() throws EWSServiceException {
		DateTime today = new DateTime();
	    DateTime startTime = today.withTime(0, 0, 0, 0);
	    DateTime endTime = today.withTime(23, 59, 59, 0);
	    
//	    System.out.println("start : " + startTime.toString("yyyy-MMM-dd HH-mm-ss"));
//	    System.out.println("end : " + endTime.toString("yyyy-MMM-dd HH-mm-ss"));

		try {
			CalendarFolder cf = CalendarFolder.bind(service,
					WellKnownFolderName.Calendar);
			FindItemsResults<Appointment> findResults = cf
					.findAppointments(new CalendarView(startTime.toDate(), endTime.toDate()));
			return findResults.getItems().size();
		} catch (Exception e) {
			e.printStackTrace();
			throw new EWSServiceException(e);
		}
	}
	
	/**
	 * get appointment from item id 
	 */
	@Override
	public Appointment getSchedule(ItemId itemId) throws EWSServiceException {
		try {
			Item itm = Item.bind(service, itemId, PropertySet.FirstClassProperties);
			Appointment appointment = Appointment.bind(service, itm.getId());
			return appointment;
		} catch (Exception e) {
			e.printStackTrace();
			throw new EWSServiceException(e);
		}
	}
	
	/**
	 * get calendar list 
	 */
	@Override
	public ArrayList<Appointment> getScheduleList(Date startTime, Date endTime, int maxItemSize) throws EWSServiceException {
		FindItemsResults<Appointment> findResults = null;
		ArrayList<Appointment> resultList = new ArrayList<Appointment>();
		try {
			CalendarFolder cf = CalendarFolder.bind(service,
					WellKnownFolderName.Calendar);
			CalendarView view = maxItemSize <= 0 ? 
					new CalendarView(startTime, endTime) : new CalendarView(startTime, endTime, maxItemSize);
			findResults = cf.findAppointments(view);
			for (Appointment appt : findResults.getItems()) {
				Appointment appointment = this.getSchedule(appt.getId());
				resultList.add(appointment);
			}
			return resultList;
		} catch (Exception e) {
			e.printStackTrace();
			throw new EWSServiceException(e);
		}
	}
	
	/**
	 * delete appoint 
	 */
	@Override
	public void deleteSchedule(Appointment appointment) throws EWSServiceException {
		try {
			appointment.delete(DeleteMode.MoveToDeletedItems);
		} catch(Exception e) {
			e.printStackTrace();
			throw new EWSServiceException(e);
		}
	}
	
	/**
	 * 일정 리스트 페이지 별 조회
	 */
	@Override
    public ArrayList<Appointment> getScheduleSearchPageList(Date startTime, Date endTime, 
    		int searchType, String searchKeyWord, int sizePerPage, int currentPage) throws EWSServiceException
    {
    	ArrayList<Appointment> resultList = new ArrayList<>();

		try {
	    	 ItemView itemview = new ItemView(sizePerPage, (currentPage - 1) * sizePerPage);
	         itemview.getOrderBy().add(AppointmentSchema.Start, SortDirection.Ascending);
	         SearchFilter.SearchFilterCollection filterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.Or);

	         // 검색 범위 조건
	         switch (searchType) {
	         case 1: //제목
	        	 filterCollection.add(new SearchFilter.ContainsSubstring(AppointmentSchema.Subject, searchKeyWord));
	        	 break;

	         case 2: //장소
	        	 filterCollection.add(new SearchFilter.ContainsSubstring(AppointmentSchema.Location, searchKeyWord));
	        	 break;

	         case 3: //이끌이
	        	 filterCollection.add(new SearchFilter.ContainsSubstring(AppointmentSchema.Organizer, searchKeyWord));
	        	 break;
	         }

	         SearchFilter.SearchFilterCollection filterCollectionDate = new SearchFilter.SearchFilterCollection(LogicalOperator.Or);

	         // 검색 조건내 일정
	         SearchFilter.SearchFilterCollection filterCollection1 = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
	         filterCollection1.add(new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, startTime));
	         filterCollection1.add(new SearchFilter.IsLessThanOrEqualTo(AppointmentSchema.End, endTime));

	         // 검색 조건보다 먼저 시작하고 검색 조건 내에서 끝나는 일정 or 검색 종료일 이후 끝나는 일정
	         SearchFilter.SearchFilterCollection filterCollection2 = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
	         filterCollection2.add(new SearchFilter.IsLessThanOrEqualTo(AppointmentSchema.Start, startTime));
	         filterCollection2.add(new SearchFilter.IsGreaterThan(AppointmentSchema.End, endTime));

	         // 검색 조건내 시간에서 시작해서 검색 종료 이후에 끝나는 일정
	         SearchFilter.SearchFilterCollection filterCollection3 = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
	         filterCollection3.add(new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, startTime));
	         filterCollection3.add(new SearchFilter.IsLessThan(AppointmentSchema.Start, endTime));

	         filterCollectionDate.add(filterCollection1);
	         filterCollectionDate.add(filterCollection2);
	         filterCollectionDate.add(filterCollection3);

	         filterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And, filterCollection, filterCollectionDate);

	         FolderId fId = new FolderId(WellKnownFolderName.Calendar);

	         service.findItems(fId, filterCollection, itemview);

	         FindItemsResults<Item> appointments = service.findItems(fId, filterCollection, itemview);
	         for (Item appt : appointments.getItems()) {
	        	 Appointment appointment = this.getSchedule(appt.getId());
	        	 resultList.add(appointment);
	         }
		}catch (Exception e) {
			e.printStackTrace();
			throw new EWSServiceException(e);
		}
    	return resultList;
    }
	
	/**
	 * 일정 등록
	 */
	@Override
    public String createSchedule(String subject, String body, Date startDate, Date endDate, int allDayEvent
        , int reminderMinutes, String location, String category, 
        List<String> toAddress, List<String> opAddress, List<String> reAddress, Recurrence recurrence, int isShare) 
        		throws EWSServiceException
    {
    	try {
    		Appointment appointment = new Appointment(service);
    		
    		appointment.setSubject(subject);
    		appointment.setBody(new MessageBody(BodyType.HTML, String.format("<html><body>%s<body></html>", body.replace("\n", "<br>"))));

    		appointment.setLocation(location);
    		if (allDayEvent > -1) appointment.setIsAllDayEvent((allDayEvent == 1) ? true : false);
    		appointment.setStart(startDate);
    		appointment.setEnd(endDate);

    		if (reminderMinutes == 0) {
    			appointment.setIsReminderSet(false);
    		}
    		else {
    			appointment.setIsReminderSet(true);
    			appointment.setReminderMinutesBeforeStart(reminderMinutes);
    		}
    		// 중요도
    		Sensitivity sensitivity = Sensitivity.Normal;
    		switch (isShare) {
    		case 0: sensitivity = Sensitivity.Private;
    		break;
    		case 1: sensitivity = Sensitivity.Normal;
    		break;
    		default: sensitivity = Sensitivity.Normal;
    		break;
    		}

    		appointment.setSensitivity(sensitivity);
    		boolean isMetting = false;

    		if (toAddress != null) {
    			for (String email : toAddress) {
    				appointment.getRequiredAttendees().add(email);
    				isMetting = true;
    			}
    		}

    		if (opAddress != null) {
    			for (String email : opAddress) {
    				appointment.getOptionalAttendees().add(email);
    				isMetting = true;
    			}
    		}

    		if (reAddress != null) {
    			for (String email : reAddress) {
    				appointment.getResources().add(email);
    				isMetting = true;
    			}
    		}

    		if (!category.isEmpty()) {
    			StringList slCategory = new StringList(Arrays.asList(category.split(",")));
    			appointment.setCategories(slCategory);
    		}

    		if (recurrence != null) {
    			appointment.setRecurrence(recurrence);
    		}

    		if (isMetting) {
    			appointment.save(SendInvitationsMode.SendToAllAndSaveCopy); // 미팅
    		}
    		else {
    			appointment.save(SendInvitationsMode.SendToNone);   // 개일 일정
    		}
    		return appointment.getId().getUniqueId();

    	}
    	catch (Exception e) {
    		e.printStackTrace();
    		throw new EWSServiceException(e); 
    	}
    }
	
	/**
	 * 일정 수정
	 */
	@Override
    public String modifyAppointment(ItemId itemId, String subject, String body, Date startDate, Date endDate, 
    		int allDayEvent, int reminderMinutes, String location, String category, 
    		List<String> toAddress, List<String> opAddress, List<String> reAddress, 
    		Recurrence recurrence, int isShare) throws EWSServiceException
    {
    	try {
    		Appointment appointment = Appointment.bind(service, itemId);

    		boolean isRecurring = false;
    		isRecurring = appointment.getIsRecurring();

    		if (isRecurring) {
    			appointment = Appointment.bindToRecurringMaster(service, itemId);
    		}

    		appointment.setSubject(subject);
    		appointment.setBody(new MessageBody(BodyType.HTML, String.format("<html><body>%s<body></html>", body.replace("\n", "<br>"))));

    		appointment.setLocation(location);
    		if (allDayEvent > -1) {
    			appointment.setIsAllDayEvent((allDayEvent == 1) ? true : false);
    		}
    		appointment.setStart(startDate);
    		appointment.setEnd(endDate);

    		if (reminderMinutes == 0) {
    			appointment.setIsReminderSet(false);
    		}
    		else {
    			appointment.setIsReminderSet(true);
    			appointment.setReminderMinutesBeforeStart(reminderMinutes);
    		}
    		// 중요도
    		Sensitivity sensitivity = Sensitivity.Normal;
    		switch (isShare) {
    		case 0: sensitivity = Sensitivity.Private;
    		break;
    		case 1: sensitivity = Sensitivity.Normal;
    		break;
    		default: sensitivity = Sensitivity.Normal;
    		break;
    		}

    		appointment.setSensitivity(sensitivity);
    		boolean isMetting = false;

    		appointment.getRequiredAttendees().clear();
    		appointment.getOptionalAttendees().clear();
    		appointment.getResources().clear();

    		if (toAddress != null) {
    			for (String email : toAddress) {
    				appointment.getRequiredAttendees().add(email);
    				isMetting = true;
    			}
    		}

    		if (opAddress != null) {
    			for (String email : opAddress) {
    				appointment.getOptionalAttendees().add(email);
    				isMetting = true;
    			}
    		}

    		if (reAddress != null) {
    			for (String email : reAddress) {
    				appointment.getResources().add(email);
    				isMetting = true;
    			}
    		}

    		if (!category.isEmpty()) {
    			StringList slCategory = new StringList(Arrays.asList(category.split(",")));
    			appointment.setCategories(slCategory);
    		}

    		if (recurrence != null) {
    			appointment.setRecurrence(recurrence);
    		}

//    		if (isMetting) {
//    			appointment.save(SendInvitationsMode.SendToAllAndSaveCopy); // 미팅
//    		}
//    		else {
//    			appointment.save(SendInvitationsMode.SendToNone);   // 개일 일정
//    		}
    		appointment.update(ConflictResolutionMode.AlwaysOverwrite);

    		return appointment.getId().getUniqueId();

    	}
    	catch (Exception e) {
    		e.printStackTrace();
    		throw new EWSServiceException(e); 
    	}
    }
	
	/**
	 * 일정 첨부 파일 로드
	 */
	@Override
    public FileAttachment getAppointmentAttachFile(ItemId itemId,  String attachId) throws EWSServiceException
    {
    	try {
    		Appointment schedule = getSchedule(itemId);
    		AttachmentCollection attachmentsCol = schedule.getAttachments();
    		for (int i = 0; i < attachmentsCol.getCount(); i++) {
    			FileAttachment attachment = (FileAttachment)attachmentsCol.getPropertyAtIndex(i);
    			if (attachment.getContentId().equals(attachId)) {
    				return attachment;
    			}
    		}
    	} 
    	catch (Exception e) {
    		e.printStackTrace();
    		throw new EWSServiceException(e);
    	}
		return null;
    }
}

