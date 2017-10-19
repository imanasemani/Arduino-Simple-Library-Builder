#include "Blinking.h"
#include <string.h>
#include <inttypes.h>
#include <Print.h>
#include <Stream.h>

Blinking::Blinking()
{

}

void Blinking::startBlink(int pin,int dlay)
{
	_pin=pin;
	pinMode(_pin,OUTPUT);

	digitalWrite(_pin,HIGH);
	delay(dlay);
	digitalWrite(_pin,LOW);
	delay(dlay);
}
