package com.aurumsoft.util;

import java.util.HashMap;

public class SoldGoods {
	private String name;
	private double sellPrice;
	/**
	 * key - day of month value - quantity
	 */
	private HashMap<Integer, Integer> quantities = new HashMap<Integer, Integer>();
	private HashMap<Integer, Double> priceRealNDS = new HashMap<Integer, Double>();
	private HashMap<Integer, Double> priceNDS = new HashMap<Integer, Double>();

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public double getSellPrice() {
		return sellPrice;
	}

	public void setSellPrice(double sellPrice) {
		this.sellPrice = sellPrice;
	}

	public HashMap<Integer, Integer> getQuantities() {
		return quantities;
	}

	public HashMap<Integer, Double> getPriceRealNDS() {
		return priceRealNDS;
	}

	public HashMap<Integer, Double> getPriceNDS() {
		return priceNDS;
	}

}
