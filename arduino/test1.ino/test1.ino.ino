int val;
int i;
byte ler;
int k; // color
//int vk;
void setup() {
  // put your setup code here, to run once:
  
  Serial.begin(9600);
  pinMode(13, INPUT);
  pinMode(3,OUTPUT); //r
  pinMode(5,OUTPUT); //g
  pinMode(6,OUTPUT); //b
}

void loop() {
  // put your main code here, to run repeatedly:
  //if (Serial.available()>0){
    if (digitalRead(13)==1){
      Serial.write(255);
    }else{
      i=analogRead(0)/4;
      val = (i>=255)?254:i;
      Serial.write(val);  
    }
  
  
    ler=Serial.read();
    if(ler<255){
      //Serial.write(ler);
      switch(ler){
        case 0:{
          k = 3;
          break;
        }
        case 1: {
          k = 5;
          break;
        }
        case 2:{
          k = 6;
          break;
        }
        default:{
          analogWrite(k,ler);
          break;
        }
      }
    }
  //}
  delay(50);


}
