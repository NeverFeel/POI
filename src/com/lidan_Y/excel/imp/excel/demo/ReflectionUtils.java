package com.lidan_Y.excel.imp.excel.demo;

/**
 * 
 * @author ilidan_Y
 *
 */
public class ReflectionUtils {

	public static String createGetMethodName(String methodName) {
		String method ="get"+methodName.substring(0,1).toUpperCase()+methodName.substring(1);
		return method;
	}

	public static String createSetMethodName(String methodName) {
		String method ="set"+methodName.substring(0,1).toUpperCase()+methodName.substring(1);
		return method;
	}

}
