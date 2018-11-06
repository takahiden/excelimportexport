// Generated File. Do Not Modify
/*
Copyright (c) 2008,2017, Oracle and/or its affiliates. All rights reserved. 

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License. 
*/

package sqldeveloper.extension.excelimportexport;

import java.awt.Image;
import javax.swing.Icon;

import oracle.dbtools.raptor.utils.MessagesBase;

public class ExtensionResources extends MessagesBase {
    // Generated Resource Keys
    public static final String SUCCESS_TITLE = "SUCCESS_TITLE"; //$NON-NLS-1$
    public static final String ERROR_NOT_CONNECT = "ERROR_NOT_CONNECT"; //$NON-NLS-1$
    public static final String NOT_DELETE_BEFORE_IMPORT = "NOT_DELETE_BEFORE_IMPORT"; //$NON-NLS-1$
    public static final String ERROR_EXPORT = "ERROR_EXPORT"; //$NON-NLS-1$
    public static final String IMPORT_TITLE = "IMPORT_TITLE"; //$NON-NLS-1$
    public static final String CONFIRM_TITLE = "CONFIRM_TITLE"; //$NON-NLS-1$
    public static final String CANCEL_EXPORT = "CANCEL_EXPORT"; //$NON-NLS-1$
    public static final String ACTION_IMPORT_LABEL = "ACTION_IMPORT_LABEL"; //$NON-NLS-1$
    public static final String FILE_NOT_EXISTS_TITLE = "FILE_NOT_EXISTS_TITLE"; //$NON-NLS-1$
    public static final String DELETE_BEFORE_IMPORT = "DELETE_BEFORE_IMPORT"; //$NON-NLS-1$
    public static final String EXPORT_TITLE = "EXPORT_TITLE"; //$NON-NLS-1$
    public static final String FILE_NOT_EXISTS = "FILE_NOT_EXISTS"; //$NON-NLS-1$
    public static final String ACTION_EXPORT_LABEL = "ACTION_EXPORT_LABEL"; //$NON-NLS-1$
    public static final String SUCCESS_IMPORT = "SUCCESS_IMPORT"; //$NON-NLS-1$
    public static final String FILE_ALREADY_EXISTS = "FILE_ALREADY_EXISTS"; //$NON-NLS-1$
    public static final String FILE_FORMAT_NAME = "FILE_FORMAT_NAME"; //$NON-NLS-1$
    public static final String LOG_OUTPUT = "LOG_OUTPUT"; //$NON-NLS-1$
    public static final String ERROR_TITLE = "ERROR_TITLE"; //$NON-NLS-1$
    public static final String ERROR_IMPORT = "ERROR_IMPORT"; //$NON-NLS-1$
    public static final String CANCEL_TITLE = "CANCEL_TITLE"; //$NON-NLS-1$
    public static final String CANCEL_IMPORT = "CANCEL_IMPORT"; //$NON-NLS-1$
    public static final String SUCCESS_EXPORT = "SUCCESS_EXPORT"; //$NON-NLS-1$
    public static final String ERROR_EXPORT_NO_QUERY = "ERROR_EXPORT_NO_QUERY"; //$NON-NLS-1$

    private static final String BUNDLE_NAME = "sqldeveloper.extension.excelimportexport.ExtensionResources"; //$NON-NLS-1$

    private static final ExtensionResources INSTANCE = new ExtensionResources();
    
    private ExtensionResources() {
        super(BUNDLE_NAME, ExtensionResources.class.getClassLoader());
    }
    
//    public static ResourceBundle getBundle() {
//        return INSTANCE.getResourceBundle();
//    }
    
//    /**
//     * @deprecated use getBundle()
//     */
//   public static ResourceBundle getInstance() {
//        return getBundle();
 //   }
    
    public static String getString( String key ) {
        return INSTANCE.getStringImpl(key);
    }
    
    public static String get( String key ) {
        return getString(key);
    }
    
    public static Image getImage( String key ) {
        return INSTANCE.getImageImpl(key);
    }
    
    public static String format(String key, Object ... arguments) {
        return INSTANCE.formatImpl(key, arguments);
    }

    public static Icon getIcon(String key) {
        return INSTANCE.getIconImpl(key);
    }
    
    public static Integer getInteger(String key) {
        return INSTANCE.getIntegerImpl(key);
    }

}
