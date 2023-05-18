`timescale 1ns / 1ps   //10^3 -> 3
 
module tb();
 
  
 
  
  reg clk16 = 0;
  reg clk8 = 0;  ///initialize variable
  reg clk_9 = 0;
  #here 9 is (MHZ )   1/9*10^6 = 111.111 *10^-9s timeperiod = 
 
   always #31.25 clk16 = ~clk16;
   always #62.5 clk8 = ~clk8;
  
  always #55.555 clk_9 = ~clk_9;
  
 
 
  initial begin
    $dumpfile("dump.vcd");
    $dumpvars;
  end
 
 
  initial begin
    #200;
    $finish();
  end
  
endmodule
