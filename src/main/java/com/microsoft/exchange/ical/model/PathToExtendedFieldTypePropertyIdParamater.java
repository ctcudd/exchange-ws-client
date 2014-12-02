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
package com.microsoft.exchange.ical.model;

import net.fortuna.ical4j.model.parameter.XParameter;
import net.fortuna.ical4j.model.property.XProperty;

import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.PathToExtendedFieldType;

/**
 * {@link XProperty} intended to store the value from {@link ExtendedPropertyType#getExtendedFieldURI()}.getPropertyId()
 * 
 * @author ctcudd
 *
 */
public class PathToExtendedFieldTypePropertyIdParamater extends XParameter {

	private static final long serialVersionUID = -8857164759801589763L;
	public static final String PATH_TO_EXTENDED_FIELD_TYPE_PROPERTY_ID = "X-EWS-PATH-TO-EXTENDED-FIELD-TYPE-PROPERTY-ID";

	public PathToExtendedFieldTypePropertyIdParamater(PathToExtendedFieldType path) {
		super(PATH_TO_EXTENDED_FIELD_TYPE_PROPERTY_ID,path.getPropertyId().toString());
	}
}
