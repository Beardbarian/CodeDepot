#Single line for reset the network stack
ipconfig /release && ipconfig /flushdns && ipconfig /renew && ipconfig /registerdns && netsh int ip reset && netsh winsock reset
