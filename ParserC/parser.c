#include <stdio.h>
#include <stdlib.h>

#define PU_AND 0;
#define PU_OR 1;
#define PU_NOT 2;
#define PU_LIST 3;
#define PU_IN 4;
#define PU_LITERAL 5;

typedef struct {
	char *name;
	int type;
	struct pu *list;
	char *chars;
	int params[2];
} parse_unit;

typedef struct pu {
	int count;
	parse_unit **items;
} parse_units;

typedef struct {
	int start;
	int end;
	int index;
	int parsed;
	struct su *list;
} parse_result;

typedef struct su {
	int count;
	parse_result **items;
} parse_results;

typedef struct {
	FILE *handle;
	int length;
	int position;
} stream_info;

parse_unit* CreateUnit() {
	return (parse_unit*) malloc(sizeof(parse_unit));
};

parse_unit* CreateAnd(parse_units *list) {
	parse_unit *unit;
	unit = CreateUnit();
	unit->type = PU_AND;
	unit->list = list;
	return unit;
};

parse_unit* CreateOr() {
	parse_unit *unit;
	unit = CreateUnit();
	unit->type = PU_OR;
	return unit;
};

parse_unit* CreateNot() {
	parse_unit *unit;
	unit = CreateUnit();
	unit->type = PU_NOT;
	return unit;
};

parse_unit* CreateList() {
	parse_unit *unit;
	unit = CreateUnit();
	unit->type = PU_LIST;
	return unit;
};

parse_unit* CreateIn(char *characters) {
	parse_unit *unit;
	unit = CreateUnit();
	unit->type = PU_IN;
	unit->chars = characters;
	unit->params[0]=strlen(characters);
	return unit;
};

parse_unit* CreateLiteral(char *characters) {
	parse_unit *unit;
	unit = CreateUnit();
	unit->type = PU_LITERAL;
	unit->chars = characters;
	unit->params[0]=strlen(characters);
	return unit;
};

stream_info m_stream;

parse_result ParseUnit(parse_unit*);

parse_result ParseAnd(parse_unit* unit) {
	parse_result result = {0,-1,0,0,NULL};
	parse_result sub_result;

	int index;

	for (index=0; index<unit->list->count; index++) {
		sub_result = ParseUnit(unit->list->items[index]);
		//printf("x%i:",sub_result.parsed);
		if (sub_result.parsed==0) {
			return result;
		};
		AppendResult(result->list, &sub_result);
	}

	result.parsed = 1;
	return result;
};


parse_result ParseOr(parse_unit* unit) {
};

parse_result ParseNot(parse_unit* unit) {
};

parse_result ParseList(parse_unit* unit) {
};

parse_result ParseIn(parse_unit* unit) {
	parse_result result = {0,-1,0,0,NULL};
	if (m_stream.position<m_stream.length) {
		char symbol = fgetc(m_stream.handle);
		int offset;
		for (offset=0; offset<unit->params[0]; offset++) {
			if (symbol==unit->chars[offset]) {
				result.index=offset+1;
				result.end = result.start = m_stream.position;
				result.parsed = 1;
				m_stream.position++;

				return result;
			};
		};
		fseek(m_stream.handle, 0, m_stream.position);
		return result;
	};

	return result;
};

parse_result ParseLiteral(parse_unit* unit) {
	parse_result result = {0,-1,0,0,NULL};
	if ((m_stream.position + unit->params[0]-1)<m_stream.length) {
		int offset;
		for (offset=0; offset<unit->params[0]; offset++) {
			char symbol = fgetc(m_stream.handle);
			if (symbol!=unit->chars[offset]) {
				fseek(m_stream.handle, 0, m_stream.position);
				return result;
			}
		};
		result.start = m_stream.position;
		m_stream.position += unit->params[0];
		result.end = m_stream.position-1;
		result.parsed = 1;
		return result;
	};

	return result;
};

parse_result ParseUnit(parse_unit* unit) {
	switch (unit->type) {
		case 0:
			return ParseAnd(unit);
		case 1:
			//return ParseOr(unit);
		case 2:
			//return ParseNot(unit);
		case 3:
			//return ParseList(unit);
		case 4:
			return ParseIn(unit);
		case 5:
			return ParseLiteral(unit);
	};

	parse_result empty={0,-1,0,0,NULL};
	return empty;
};

void AppendUnit(parse_units *the_array, parse_unit *unit) {
	the_array->count++;
	the_array->items = (parse_unit**) realloc(the_array->items, the_array->count*sizeof(parse_unit*));
	the_array->items[the_array->count-1] = unit;
};

void AppendResult(parse_results *the_array, parse_result *result) {
	the_array->count++;
	the_array->items = (parse_result**) realloc(the_array->items, the_array->count*sizeof(parse_result*));
	the_array->items[the_array->count-1] = result;
};

stream_info OpenStream(char *filename) {
	stream_info stream;

	stream.handle=NULL;
	stream.position=0;
	stream.length=0;

	if ((stream.handle=fopen(filename, "r"))==NULL) {
		printf("file not open");
	} else {
		fseek(stream.handle,0,SEEK_END);
		stream.length = ftell(stream.handle);
		fseek(stream.handle,0,SEEK_SET);
	}

	return stream;
};

int main(int argc, char **argv) {
	parse_unit *cat=CreateIn("abc");
	parse_unit *dog=CreateLiteral("dog");
	parse_unit *cow=CreateLiteral("cow");

	parse_units and_array = {0, NULL};
	parse_result result;

	m_stream = OpenStream("test.txt");

	AppendUnit(&and_array, dog);
	AppendUnit(&and_array, cow);
	//AppendUnit(&and_array, cow);
	parse_unit *joined = CreateAnd(&and_array);

	result = ParseUnit(joined);

	printf("%i", result.parsed);

	fclose(m_stream.handle);
	
	/*

	parse_result result = {0,0,0,NULL};

	result = ParseIn2(cat);
	printf("%i", result.end);



	AppendUnit(&and_array, cat);
	AppendUnit(&and_array, dog);
	AppendUnit(&and_array, cow);
	parse_unit *joined = CreateAnd(&and_array);

	int parsed = ParseUnit(joined, &result);
*/

//	int parsed = ParseUnit(cat, &result);
//	printf("%i", parsed);


	return 0;
};

	//int parsed = ParseIn(cat, &result);
	//printf("%i", result.index);
	//int parsed = ParseLiteral(dog, &result);
	//printf("%i", parsed);

	//char symbol = fgetc(m_stream.handle);
	//printf("%c", symbol);

//	parse_units the_array;
//	the_array.items = NULL;
//	the_array.count=0;

	//printf("%i", m_stream.length);
	
	/*
	AppendUnit(&the_array, unit);
	unit = CreateNot();
	AppendUnit(&the_array, unit);
	*/
	//printf("%i\n", the_array.items[1]->type);
	//printf("%s", unit->chars);