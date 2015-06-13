package test;


import org.testng.Assert;
import org.testng.annotations.Test;

public class StringOrderTest {

	@Test
	public void testOrder(){
		String one = "一";
		String three = "a";
		String five = "b";
		String seven = "c";
		
		System.out.println(three.compareTo(five));
		System.out.println(five.compareTo(three));
		System.out.println(five.compareTo(seven));
		Assert.assertTrue(three.compareTo(five) < 0);
		Assert.assertTrue(seven.compareTo(five) > 0);
		
	}
}
