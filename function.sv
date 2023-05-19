module func();
  
  function automatic [] f(ref bit [7:0] arr[8]);
    
    return f
  endfunction
endmodule


//creating own constructor
class first;
  
  int data;
  
  function new(input int datain = 0);
  data = datain;
  endfunction
  
  
endclass
 
 
module tb;
  
  first f1;
  
  initial begin
    f1 = new(23);
    $display("Data : %0d", f1.data); 
  end
  
  
endmodule



////handling multiple data members in constructor

class first;
  
  int data1;
  bit [7:0] data2;
  shortint data3;
  
  function new(input int data1 = 0, input bit[7:0] data2 = 8'h00, input shortint data3 = 0);
   this.data1 = data1;
   this.data2 = data2;
   this.data3 = data3;    
  endfunction
  
  
endclass
 
 
module tb;
  
  first f1;
  
  initial begin
    //f1 = new(23,,35); ///follow position
    f1 = new( .data2(4), .data3(5), .data1(23)); //follow name
    $display("Data1 : %0d, Data2 : %0d and Data3 : %0d", f1.data1, f1.data2, f1.data3); 
  end
  
  
endmodule


/// Asssignment 

class Second;
  bit [7:0] a, b, c;
  function new(input bit [2:0] a,b, c);
    this.a = a;
    this.b = b;
    this.c = c;
  endfunction
endclass

module check;
  Second obj;
  initial begin
    obj = new(.a(2),.b(4),.c(56));
    
    $display("%0d %0d %0d",a, b,c);
  end
endmodule
