package cn.wangkf.util;

import com.google.common.collect.Maps;
import com.to8to.sc.compatible.RPCException;
import lombok.extern.slf4j.Slf4j;

import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

/**
 * Created by stanley.wang on 2020/4/16.
 */
@Slf4j
public class ObjectCompareUtil {

    private static final String NEW_VALUE_MARK = "newValue";
    private static final String OLD_VALUE_MARK = "oldValue";

    public static boolean compareObject(Object oldObject, Object newObject) {
        return compareObject(oldObject, newObject, null);
    }

    /**
     * 比较两个object的属性值是否全部相等
     * @param oldObject
     * @param newObject
     * @param childFieldClass
     * @return
     */
    public static boolean compareObject(Object oldObject, Object newObject, List<Class> childFieldClass) {
        Map<String, Map<String, Object>> resMap = Maps.newHashMap();
        compareFields(resMap, null, oldObject, newObject, childFieldClass);
        log.info("resMap={}", resMap);
        return resMap.size() > 0;
    }

    /**
     * 递归比较对象的属性值是否相等
     * @param resMap 性值是不相等的MAP
     * @param fatherFieldName 对象属性名称
     * @param oldObject 旧对象
     * @param newObject 新对象
     * @param childFieldClass 是否要比较属性值object的list<class>
     */
    public static void compareFields(Map<String, Map<String, Object>> resMap, String fatherFieldName,
                                     Object oldObject, Object newObject, List<Class> childFieldClass) {

        try{
            // 只有两个对象都是同一类型的才有可比性
            if (oldObject.getClass() != newObject.getClass()) {
                return;
            }

            Class clazz = oldObject.getClass();
            //获取object的所有属性
            PropertyDescriptor[] pds = Introspector.getBeanInfo(clazz, Object.class).getPropertyDescriptors();
            for (PropertyDescriptor pd : pds) {
                //遍历获取属性名
                String fieldName = pd.getName();
                String fullFieldName = fatherFieldName != null ?
                        fatherFieldName + "_" + fieldName : fieldName;

                //获取属性的get方法
                Method readMethod = pd.getReadMethod();

                // 在oldObject上调用get方法等同于获得oldObject的属性值
                Object oldValue = readMethod.invoke(oldObject);
                // 在newObject上调用get方法等同于获得newObject的属性值
                Object newValue = readMethod.invoke(newObject);
                // 适用于 updateByPrimaryKeySelective
                if (newValue == null || newValue instanceof List || newValue instanceof Map) {
                    continue;
                }

                // 原值为null，新值不为null
                if(oldValue == null && newValue != null){
                    Map<String,Object> valueMap = Maps.newHashMap();
                    valueMap.put(OLD_VALUE_MARK, oldValue);
                    valueMap.put(NEW_VALUE_MARK, newValue);
                    resMap.put(fullFieldName, valueMap);
                    continue;
                }

                // 比较该属性object的值是否相等
                if (isHaveChirldObject(childFieldClass, newValue)) {
                    compareFields(resMap, fullFieldName, oldValue, newValue, childFieldClass);
                    continue;
                }

                // 比较这两个值是否相等,不等就可以放入map了
                if (!oldValue.equals(newValue)) {
                    Map<String,Object> valueMap = Maps.newHashMap();
                    valueMap.put(OLD_VALUE_MARK, oldValue);
                    valueMap.put(NEW_VALUE_MARK, newValue);
                    resMap.put(fullFieldName, valueMap);
                }
            }
        } catch (Exception e){
            log.error("e={}", e);
            throw new RPCException(e);
        }
    }

    /**
     * 比较该属性值是否在要比较的指定class里面
     * @param childFieldClass
     * @param value
     * @return
     */
    private static boolean isHaveChirldObject(List<Class> childFieldClass, Object value) {
        if (childFieldClass == null || childFieldClass.size() == 0) {
            return false;
        }
        for (Class c : childFieldClass) {
            if (c == value.getClass()) {
                return true;
            }
        }
        return false;
    }

}
