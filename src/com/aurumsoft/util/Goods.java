package com.aurumsoft.util;

import java.math.BigDecimal;
import java.util.TreeSet;

public class Goods implements Comparable<Goods> {

	private String name;

	private String category;
	private int statementRowIndex;
	public static final int SELL_PRICE_ROUND = 1;
	public static final int BUY_PRICE_ROUND = 5;
	private TreeSet<GoodsPrice> priceList = new TreeSet<>();

	public Goods(String name, String category) {
		this.name = name;
		this.category = category;
	}

	public void addPriceElement(int quantity, double priceWVAT,
			int wharehouseRowIndex) {
		priceList.add(new GoodsPrice(quantity, priceWVAT, wharehouseRowIndex));
	}

	public TreeSet<GoodsPrice> getPriceList() {
		return priceList;
	}

	public String getCategory() {
		return category;
	}

	public void decreaseQuantity() {
		for (GoodsPrice priceElement : priceList) {
			if (priceElement.getQuantity() > 0) {
				priceElement.decreaseQuantity();
				break;
			}
		}
	}

	public String getName() {
		return name;
	}

	public int getStatementRowIndex() {
		return statementRowIndex;
	}

	public void setStatementRowIndex(int statementRowIndex) {
		this.statementRowIndex = statementRowIndex;
	}

	public static double round(double unrounded) {
		return round(unrounded, 2, BigDecimal.ROUND_HALF_UP);
	}

	public static double round(double unrounded, int precision) {
		return round(unrounded, precision, BigDecimal.ROUND_HALF_UP);
	}

	public static double round(double unrounded, int precision, int roundingMode) {
		BigDecimal bd = new BigDecimal(unrounded);
		BigDecimal rounded = bd.setScale(precision, roundingMode);
		return rounded.doubleValue();
	}

	public int getTotalQuantity() {
		int totalQuantity = 0;
		for (GoodsPrice priceElement : priceList) {
			totalQuantity += priceElement.getQuantity();
		}
		return totalQuantity;
	}

	@Override
	public String toString() {
		return name + " [category=" + category + "]";
	}

	@Override
	public int compareTo(Goods o) {
		return o.getTotalQuantity() - getTotalQuantity();
	}

	@Override
	public boolean equals(Object obj) {
		if (obj == null) {
			return false;
		}
		if (obj instanceof Goods) {
			Goods gObj = (Goods) obj;
			if (this.name.equals(gObj.name) && this.category == gObj.category) {
				return true;
			} else {
				return false;
			}
		} else {
			return false;
		}

	}

}
