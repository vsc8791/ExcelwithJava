package dao;

import entities.*;

import java.text.SimpleDateFormat;
import java.util.*;;

public class ProductModel {

	public List<Product> findAll() {
		try {
			
			SimpleDateFormat sdf =new SimpleDateFormat("yyyy-MM-dd"); 
			List<Product> result = new ArrayList<Product>();
			result.add(new Product("P1", "Laptops", 35000, 4, sdf.parse("2017-01-02")));
			result.add(new Product("P2", "Mobiles", 16990, 5, sdf.parse("2016-10-20")));
			result.add(new Product("P3", "Fridges", 28000, 3, sdf.parse("2016-07-08")));
			result.add(new Product("P4", "Theatre", 69000, 2, sdf.parse("2016-11-28")));

			return result;
		} catch (Exception ex) {

			return null;
		}
	}

}
