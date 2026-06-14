#include <ApplicationServices/ApplicationServices.h>
#include <stdlib.h>
#include <unistd.h>

int main(int argc, char *argv[]) {
    int delta = argc > 1 ? atoi(argv[1]) : -1;
    int repeats = argc > 2 ? atoi(argv[2]) : 1;
    if (repeats < 1) repeats = 1;

    for (int i = 0; i < repeats; i++) {
        CGEventRef event = CGEventCreateScrollWheelEvent(
            NULL,
            kCGScrollEventUnitLine,
            1,
            delta
        );
        if (event == NULL) return 1;
        CGEventPost(kCGHIDEventTap, event);
        CFRelease(event);
        usleep(120000);
    }

    return 0;
}
