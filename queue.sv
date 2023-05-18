module queue;
  int arr[$];
  int j =0;
  initial begin
    arr = {1,2,3};
    $display("%0p",arr);
    arr.push_front(7); // 7 1 2 3
    
    arr.push_back(9); // 7 1 2 3 9
    
    arr.insert(2,10) // at index 2, put value 10
    
    j = arr.pop_front();
    
    $display("%0d",j);
    k = arr.pop_back();
    
    arr.delete(2);  // deleting the value from the array at the index 2
    
    ar
  end
endmodule
