c
c --- Calculate Matrix Values ---
c
      mat(1,1)=j10y+j30y+m10*p**2+18.0*m30*p**2+18.0*m30*p**2*cos(q2)*
     . cos(q3)-j30y*sin(q3)**2+j30z*sin(q3)**2-9.0*m30*p**2*sin(q3)**2
      mat(1,2)=j30y+9.0*m30*p**2+9.0*m30*p**2*cos(q2)*cos(q3)-j30y*sin(
     . q3)**2+j30z*sin(q3)**2-9.0*m30*p**2*sin(q3)**2
      mat(1,3)=-9.0*m30*p**2*sin(q2)*sin(q3)
      mat(2,2)=j30y+9.0*m30*p**2-j30y*sin(q3)**2+j30z*sin(q3)**2-9.0*m30
     . *p**2*sin(q3)**2
      mat(2,3)=0
      mat(3,3)=j30x+9.0*m30*p**2
