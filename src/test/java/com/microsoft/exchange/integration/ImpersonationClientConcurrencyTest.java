/**
 * See the NOTICE file distributed with this work
 * for additional information regarding copyright ownership.
 * Board of Regents of the University of Wisconsin System
 * licenses this file to you under the Apache License,
 * Version 2.0 (the "License"); you may not use this file
 * except in compliance with the License. You may obtain a
 * copy of the License at:
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on
 * an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
/**
 * 
 */
package com.microsoft.exchange.integration;

import java.util.Date;
import java.util.HashSet;
import java.util.Set;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import junit.framework.Assert;

import org.apache.commons.lang.time.DateUtils;
import org.apache.commons.lang.time.StopWatch;
import org.apache.commons.math.stat.descriptive.SynchronizedSummaryStatistics;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.microsoft.exchange.ExchangeDateUtils;
import com.microsoft.exchange.config.ImpersonationConfig;
import com.microsoft.exchange.exception.ExchangeRuntimeException;
import com.microsoft.exchange.impl.ExchangeCalendarDataDao;
import com.microsoft.exchange.impl.ThreadLocalImpersonationConnectingSIDSourceImpl;
import com.microsoft.exchange.messages.CreateItem;
import com.microsoft.exchange.messages.CreateItemResponse;
import com.microsoft.exchange.messages.FindItem;
import com.microsoft.exchange.messages.FindItemResponse;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.ConnectingSIDType;
import com.microsoft.exchange.types.ItemIdType;

/**
 * Perform some tests targeted at observing throttling policy and other issues
 * when using a number of concurrent connections configured with impersonation support.
 * 
 * @author Nicholas Blair
 */
@RunWith(SpringJUnit4ClassRunner.class)
//@ContextConfiguration(locations= {"classpath:test-contexts/exchangeContext.xml"})
@ContextConfiguration(classes=ImpersonationConfig.class)
public class ImpersonationClientConcurrencyTest extends AbstractIntegrationTest {

	@Autowired
	protected ExchangeCalendarDataDao exchangeCalendarDataDao;
	
	private int targetConcurrency;
	/**
	 * @return the targetConcurrency
	 */
	public int getTargetConcurrency() {
		return targetConcurrency;
	}
	/**
	 * @param targetConcurrency the targetConcurrency to set
	 */
	@Value("${http.maxTotalConnections}")
	public void setTargetConcurrency(int targetConcurrency) {
		this.targetConcurrency = targetConcurrency;
	}

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.integration.AbstractIntegrationTest#initializeCredentials()
	 */
	@Override
	public void initializeCredentials() {
		ConnectingSIDType connectingSID = new ConnectingSIDType();
		connectingSID.setPrincipalName(emailAddress);
		ThreadLocalImpersonationConnectingSIDSourceImpl.setConnectingSID(connectingSID);
	}

	@Test @Override
	public void getPrimaryCalendarFolder() {
		super.getPrimaryCalendarFolder();
	}
	
	@Test
	public void testConcurrentCreateItems() throws InterruptedException{
		int threadCount = 10;
		final int batchesPerThread = 6;
		final int itemsPerBatch = 10;
		
		// setup a latch to stall all threads until ready to run all at once (-1 so the last thread's run invocation triggers the start)
		final CountDownLatch startLatch = new CountDownLatch(threadCount);
		final CountDownLatch endLatch = new CountDownLatch(threadCount);
		ExecutorService executor = Executors.newFixedThreadPool(threadCount);
		
		final SynchronizedSummaryStatistics stats = new SynchronizedSummaryStatistics();
		final Set<ItemIdType> createdItemIds = new HashSet<ItemIdType>();
		try {
			for(int i = 0; i < threadCount; i++) {
				final int index = i;
				executor.submit(new Runnable() {
					@Override
					public void run() {
						try {
							//batch of CreateItem requests
							Set<CreateItem> requests = new HashSet<CreateItem>();
							for(int j = 0; j < batchesPerThread; j++) {
								//collection of items for creation
								Set<CalendarItemType> items = new HashSet<CalendarItemType>();
								for(int k =0; k < itemsPerBatch; k++){
									CalendarItemType c = new CalendarItemType();
									c.setSubject(Thread.currentThread().getName() + " thread="+index+", batch="+j+", item="+k);
									items.add(c);									
								}
								CreateItem request = requestFactory.constructCreateCalendarItem(items);
								requests.add(request);
							}
							
							startLatch.countDown();
							try {
								startLatch.await();
							} catch (InterruptedException e) {
								throw new IllegalStateException("interrupted while waiting to start", e);
							}
							
							for(CreateItem request : requests){
								StopWatch time = new StopWatch();
								time.start();
								initializeCredentials();
								Set<ItemIdType> parsed = new HashSet<ItemIdType>();
								try{
									CreateItemResponse response = ewsClient.createItem(request);
									parsed = responseUtils.getCreatedItemIds(response);
									createdItemIds.addAll(parsed);
								}catch(ExchangeRuntimeException e){
									log.error(e.getMessage());
								}finally{
									time.stop();
								}
								
								log.info(Thread.currentThread().getName() + " created " + parsed.size()+ " in " +time);
								stats.addValue(time.getTime());
							}
							
						} finally {
							endLatch.countDown();
						}
					}
				});
			}
			// now block until everybody is done
			endLatch.await();
			log.info("testConcurrentCreateItems complete for " + targetConcurrency + " threads, stats: " + stats);
		} finally {
			executor.shutdown();
		}
		int expectedCount = threadCount * batchesPerThread * itemsPerBatch;
		int actualCount = createdItemIds.size();
		
		log.info("Successfully created "+actualCount+" of "+expectedCount+" calendarItems.");
		//delete createdItemIds
		StopWatch time = new StopWatch();
		time.start();
		
		//boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItems(emailAddress, createdItemIds);
		boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItemsWithRetry(emailAddress, createdItemIds);
		time.stop();
		log.info((deleteSuccess ? "Successfully deleted ": "Failed to delete ")+actualCount+" created calendar items in "+time);
		
		//items in should equal items out
		Assert.assertEquals(expectedCount, actualCount);
	}
	
	/**
	 * 
	 * @throws InterruptedException 
	 */
	@Test
	public void testConcurrentFindItems() throws InterruptedException {
		final int threadCount = 100;
		// setup a latch to stall all threads until ready to run all at once (-1 so the last thread's run invocation triggers the start)
		final CountDownLatch startLatch = new CountDownLatch(threadCount);
		final CountDownLatch endLatch = new CountDownLatch(threadCount);
		ExecutorService executor = Executors.newFixedThreadPool(threadCount);
		final Date start = ExchangeDateUtils.makeDate(startDate);
		final Date end = ExchangeDateUtils.makeDate(endDate);
		final SynchronizedSummaryStatistics stats = new SynchronizedSummaryStatistics();
		try {
			for(int i = 0; i < threadCount; i++) {
				final int index = i;
				executor.submit(new Runnable() {
					@Override
					public void run() {
						try {
							initializeCredentials();
							Date itemStart = DateUtils.addDays(start, index);
							Date itemEnd = DateUtils.addDays(end, index);
							FindItem request = constructFindItemRequest(itemStart, itemEnd, emailAddress);
							startLatch.countDown();
							try {
								startLatch.await();
							} catch (InterruptedException e) {
								throw new IllegalStateException("interrupted while waiting to start", e);
							}
							for(int j = 0; j < 10; j++) {
								StopWatch time = new StopWatch();
								time.start();
								FindItemResponse response = null;
								
								response = ewsClient.findItem(request);
								Set<CalendarItemType> found = responseUtils.parseFindCalendarItemResponse(response);
								
								time.stop();
								//String capture = capture(response);
								//log.info(Thread.currentThread().getName() + " response: " + capture);
								log.info(Thread.currentThread().getName() + " found " + found.size()+ " events between "+itemStart.toString()+ " and "+itemEnd.toString());
								stats.addValue(time.getTime());
							}
						} finally {
							endLatch.countDown();
						}
					}
				});
			}
			// now block until everybody is done
			endLatch.await();
			log.info("testConcurrentFindItems complete for " + targetConcurrency + " threads, stats: " + stats);
		} finally {
			executor.shutdown();
		}

	}
}
