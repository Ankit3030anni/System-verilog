#Aligning edges of the different clocks

module A_C():
  reg clk_20;
  reg clk_40;
  reg clk_10;
  initial begin 
    clk_20 = 0;
    clk_10 =0;
    clk_40 = 0;
  end
  always #5 clk_10 = ~clk_10;
  always begin 
    #5;
    clk_20 = 1;
    #10;
    clk_20 = 0;
    #5;
  
  end
  always begin 
    #5;
    clk_40 = 1;
    #20;
    clk_40 = 0;
    #15;
  
  end
  
  initial begin
    #200;
    $finish();
  end
endmodule
