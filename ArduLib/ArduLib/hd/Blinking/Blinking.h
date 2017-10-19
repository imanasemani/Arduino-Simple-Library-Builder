#ifndef Blinking_h
#define Blinking_h

// arduino versioning
#if (ARDUINO >= 100)
  #include "Arduino.h"
#else
  #if defined(__AVR__)
    #include <avr/io.h>
  #endif
  #include "WProgram.h"
#endif

#include <inttypes.h>
#include <Stream.h>
#include <SoftwareSerial.h>

class Blinking
{
 public:
	Blinking();

	void startBlink(int pin, int dlay);
 private:
	int _pin;
};

#endif