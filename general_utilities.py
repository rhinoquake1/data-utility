from msilib.schema import Error
import pandas as pd

class Utility:
    def __init__(self) -> None:
        self.select_method()

    def select_method(self):
        while True:
            selections = {
                'ttr': 'Text to rows',
                'exit': 'Quit'
            }
            print('~~~~~~~~~~')
            print(selections)
            selection = input('Selection: ').lower()
            if selection == 'ttr':
                self.ttr().main()
            elif selection == 'exit':
                break
            else:
                print('Unrecognised :(')
                return

    class ttr:
        def __init__(self) -> None:
            pass

        def main(self):
            df = pd.read_clipboard()
            rerun = ''
            while rerun != 'y':
                print(df)
                column = input('Column: ')
                delimiter = input('Delimiter: ')
                if column.lower() == '!exit':
                    break
                try:
                    result = self.text_to_rows(df, column, delimiter)
                except KeyError:
                    print(">>>> Header doesn't check out")
                    continue
                print(result)
                rerun = input('Accept? (y): ').lower()
            result.to_clipboard()
            print('Copied to clipboard')

        def text_to_rows(self, df : pd.DataFrame, column, delimiter):
            s = df[column].str.split(delimiter).apply(pd.Series, 1).stack()
            s.index = s.index.droplevel(-1) # to line up with df's index
            s.name = column # needs a name to join
            # del df[column]
            df = df.drop(column, axis=1)
            df = df.join(s)
            return df



if __name__ == '__main__':
    Utility()