// tStringBuffer.h: interface for the tStringBuffer class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_TSTRINGBUFFER_H__35A7FE20_7CD4_11D4_B881_0000B45D7541__INCLUDED_)
#define AFX_TSTRINGBUFFER_H__35A7FE20_7CD4_11D4_B881_0000B45D7541__INCLUDED_

class tStringBuffer  
{
public:
  const char * getBuffer(void);
  void copyToBuffer(const char *source);
  void copyToBuffer(const char *source, size_t length);
  tStringBuffer();
  tStringBuffer(const char* source);
  tStringBuffer(const tStringBuffer& copy);
  tStringBuffer& operator=(const tStringBuffer& other);
  operator const char*();
  virtual ~tStringBuffer();


protected:
  size_t size;
  char * buffer;
};

#endif // !defined(AFX_TSTRINGBUFFER_H__35A7FE20_7CD4_11D4_B881_0000B45D7541__INCLUDED_)
