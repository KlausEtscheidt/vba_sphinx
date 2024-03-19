import sphinx.cmd.build as sb
import sys

def buildit(builder,srcdir='.'):
    # builder = 'xml'
    # builder = 'html'
    myargs = ['-M', builder, srcdir, srcdir+'/build','-E']
    # sb.build_main(myargs)
    sb.main(myargs)

if __name__ == '__main__':
    
    if len(sys.argv)<2:
        raise SystemExit('Keine Argumente übergeben')

    buildit(sys.argv[1]) 