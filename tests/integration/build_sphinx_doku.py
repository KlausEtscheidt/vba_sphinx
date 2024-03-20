import sys
import sphinx.cmd.build as sb

def buildit(builder,srcdir='.'):
    # builder = 'xml'
    # builder = 'html'
    myargs = ['-M', builder, srcdir, srcdir+'/build','-E']
    sb.main(myargs)

if __name__ == '__main__':
    if len(sys.argv)<2:
        raise SystemExit('Keine Argumente übergeben')

    buildit(sys.argv[1]) 