package com.aurumsoft.util;

import java.math.BigDecimal;

public class GoodsPrice implements Comparable<GoodsPrice> {

	private double sellPrice;
	private double buyPriceWaVAT;
	private double priceWVAT;
	private int wharehouseRowIndex;
	private int quantity;

	public GoodsPrice(int quantity, double priceWVAT, int wharehouseRowIndex) {
		this.quantity = quantity;
		this.priceWVAT = priceWVAT;
		this.wharehouseRowIndex = wharehouseRowIndex;
		if (quantity > 0) {
			this.sellPrice = Goods.round(priceWVAT / quantity * 1.5,
					Goods.SELL_PRICE_ROUND, BigDecimal.ROUND_HALF_UP);
			this.buyPriceWaVAT = Goods.round((priceWVAT / quantity) / 1.2,
					Goods.BUY_PRICE_ROUND, BigDecimal.ROUND_HALF_UP);
		}
	}

	public void decreaseQuantity() {
		priceWVAT = priceWVAT - (priceWVAT / quantity);
		quantity--;
	}

	public double getSellPrice() {
		return sellPrice;
	}

	public double getBuyPriceWaVAT() {
		return buyPriceWaVAT;
	}

	public int getWharehouseRowIndex() {
		return wharehouseRowIndex;
	}

	public double getPriceWVAT() {
		return priceWVAT;
	}

	public int getQuantity() {
		return quantity;
	}

	@Override
	public int compareTo(GoodsPrice arg0) {
		return wharehouseRowIndex - arg0.wharehouseRowIndex;
	}
}
