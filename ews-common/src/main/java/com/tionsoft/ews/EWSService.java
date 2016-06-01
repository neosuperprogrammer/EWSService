package com.tionsoft.ews;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.property.complex.recurrence.pattern.Recurrence;

/**
 * MS Exchange Web Service Interface
 * @author 남상욱
 * @since 2016.03.24
 * @brief MS Exchange Web Service Interface 
 */

public interface EWSService {
	
	/**
	 * 오늘 일정 개수
	 * @return 오늘 일정 개수
	 * @throws EWSServiceException
	 */
	public int getTodayScheduleCount() throws EWSServiceException;

	/**
	 * 일정 조회
	 * @param itemId 일정 아이디 
	 * @return Appointment 일정
	 * @throws EWSServiceException
	 */
	public Appointment getSchedule(ItemId itemId) throws EWSServiceException;
	
	/**
	 * 일정 리스트 조회 
	 * @param startTime 시작일자 
	 * @param endTime 종료일자 
	 * @param maxItemSize 최대 아이템 갯수 (0 이거나 0 보다 작으면 전체 조회)
	 * @return ArrayList<Appointment> 일정 리스트 
	 * @throws EWSServiceException 
	 */
	public ArrayList<Appointment> getScheduleList(Date startTime, Date endTime, int maxItemSize) throws EWSServiceException;
	
	/**
	 * 일정삭제 
	 * @param appointment
	 * @throws EWSServiceException
	 */
	public void deleteSchedule(Appointment appointment) throws EWSServiceException;
	
	/**
	 * 일정 리스트 페이지 별 조회
	 * @param startTime 시작일자 
	 * @param endTime 종료일자 
	 * @param searchType 검색 조건 (1:제목,2:장소,3:이끌이)
	 * @param searchKeyWord 검색어 
	 * @param sizePerPage 페이지 당 갯수 
	 * @param currentPage 현재 페이지 
	 * @return 일정 리스트 
	 * @throws EWSServiceException 
	 */
    public ArrayList<Appointment> getScheduleSearchPageList(Date startTime, Date endTime, 
    		int searchType, String searchKeyWord, int sizePerPage, int currentPage) throws EWSServiceException;
    
	/**
	 * 일정 등록
	 * @param subject
	 * @param body
	 * @param startDate
	 * @param endDate
	 * @param allDayEvent
	 * @param reminderMinutes
	 * @param location
	 * @param category
	 * @param toAddress
	 * @param opAddress
	 * @param reAddress
	 * @param recurrence
	 * @param isShare
	 * @return
	 * @throws EWSServiceException 
	 */
    public String createSchedule(String subject, String body, Date startDate, Date endDate, int allDayEvent
        , int reminderMinutes, String location, String category, 
        List<String> toAddress, List<String> opAddress, List<String> reAddress, Recurrence recurrence, int isShare) 
        		throws EWSServiceException;
    
	/**
	 * 일정 수정
	 * @param itemId
	 * @param subject
	 * @param body
	 * @param startDate
	 * @param endDate
	 * @param allDayEvent
	 * @param reminderMinutes
	 * @param location
	 * @param category
	 * @param toAddress
	 * @param opAddress
	 * @param reAddress
	 * @param recurrence
	 * @param isShare
	 * @return 등록 ID
	 * @throws EWSServiceException
	 */
    public String modifyAppointment(ItemId itemId, String subject, String body, Date startDate, Date endDate, 
    		int allDayEvent, int reminderMinutes, String location, String category, 
    		List<String> toAddress, List<String> opAddress, List<String> reAddress, 
    		Recurrence recurrence, int isShare) throws EWSServiceException;
    
    /**
	 * 일정 첨부 파일 로드
	 * @param itemId 일정 아이디 
	 * @param attachId 첨부 파일 아이디 
	 * @return 첨부 파일 
	 * @throws EWSServiceException
	 */
    public FileAttachment getAppointmentAttachFile(ItemId itemId,  String attachId) throws EWSServiceException;
    
}
