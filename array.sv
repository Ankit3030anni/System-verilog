module tb;
  
  int arr[10];///0-9
  int i =0;
  
  
  int brr[];
  
  initial begin
    $display($size(brr)); // output -> 0;
  end
  
  int arr1[5] = `{5{1}}; // output -> 1 1 1 1 1
  int arr2[5] = `{1,2,3,4,5} // adding unique value
  int arr3[5] = `{default : 4};  // 4 4 4 4 4
  
  //dynamic array
  
  int arr[];
  
  initial begin
    arr = new[5];
    for (int i =0; i <5 ; i++) begin
      arr[i] = 1;
      
    end
    arr = new[30](arr);  // updating the size ofthe array keeeping the first 5 as of there present
    foreach (arr[j]) begin
      arr[j] = j;
    end
  end
  /*
  initial begin
    
    for(i= 0; i< 10; i++) begin
      arr[i] = i;    
    end
    
    
    $display("arr : %0p", arr);
    
    
  end
  
  */
  
  /*
  initial begin
    
  foreach(arr[j]) begin //0---9
    arr[j] = 5;
    $display("%0d", arr[j]);
  end
    
  end
  */
  
  initial begin
    
    repeat(10) begin
      arr[i] = i;
      i++;
    end
    
    $display("arr : %0p",arr);
    
  end
  
  
  
endmodule
