#import <Foundation/Foundation.h>
#import <ServiceManagement/ServiceManagement.h>
//
// officekick-register-login-item [--check]
//
// Registers the CONTAINING app bundle (OfficeKickServer.app) as a login item
// through SMAppService. macOS then attributes the background item - the
// Settings "Login Items" row and the one-time "can run in the background"
// notification - to "OfficeKickServer". The legacy System Events AppleScript
// path (tell app "System Events" to make login item) gets attributed to
// "System Events.app" instead and fires that notification again on every
// re-registration, which reads like a scare prompt.
//
// --check : exit 0 iff the service is already enabled (prints STATUS <n>).
// register: exit 0 if already enabled or registration succeeded.
//
// Build: clang -framework Foundation -framework ServiceManagement \
//          -o officekick-register-login-item officekick-register-login-item.m
// The bundle must be re-signed (codesign -s -) AFTER this binary is placed in
// Contents/MacOS, or macOS invalidates the item and drops the registration.
#include <stdio.h>

int main(int argc, const char * const argv[]) {
  @autoreleasepool {
    SMAppService *svc = [SMAppService mainAppService];
    NSString *mode = (argc > 1) ? [NSString stringWithUTF8String:argv[1]] : @"register";
    SMAppServiceStatus st = [svc status];
    if ([mode isEqualToString:@"--check"]) {
      fprintf(stdout, "STATUS %ld\n", (long)st);
      return st == SMAppServiceStatusEnabled ? 0 : 1;
    }
    if (st == SMAppServiceStatusEnabled) { fprintf(stdout, "ALREADY\n"); return 0; }
    NSError *err = nil;
    if ([svc registerAndReturnError:&err]) { fprintf(stdout, "REGISTERED\n"); return 0; }
    const char *reason = err ? (err.localizedDescription.UTF8String ?: (err.domain.UTF8String ?: "unknown error")) : "unknown error";
    fprintf(stderr, "ERR %s\n", reason);
    return 1;
  }
}