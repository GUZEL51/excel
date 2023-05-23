package com.qa.test;

public class Main {

	public static void main(String[] args) {
		
		Shape s;
		
		if(Math.random()<.5) {
			s= new Circle();
		}
		else {
				s= new Square();
			}
			s.draw();
			}
		}

